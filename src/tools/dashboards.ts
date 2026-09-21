// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { Widget } from "azure-devops-node-api/interfaces/DashboardInterfaces.js";
import { z } from "zod";
import { elicitProject, elicitTeam } from "../shared/elicitations.js";

const DASHBOARD_TOOLS = {
  dashboard: "dashboard",
  dashboard_write: "dashboard_write",
};

function configureDashboardTools(server: McpServer, connectionProvider: () => Promise<WebApi>) {
  // ─── dashboard (read-only) ────────────────────────────────────────────
  server.tool(
    DASHBOARD_TOOLS.dashboard,
    "Read Azure DevOps team dashboards. Use the action parameter to specify the operation. Requires PAT authentication with the 'Dashboards: Read and manage' scope — this domain does not work with interactive OAuth (confirmed: TF400813 on every call).",
    {
      action: z.enum(["list", "get"]).describe("The action to perform. Options: list (list dashboards under a project/team), get (get full details of one dashboard, including its widgets)."),
      project: z.string().optional().describe("The name or ID of the Azure DevOps project. Reuse from prior context if already known. If not provided, a project selection prompt will be shown."),
      team: z.string().optional().describe("The name or ID of the Azure DevOps team. Reuse from prior context if already known. If not provided, a team selection prompt will be shown."),
      dashboardId: z.string().optional().describe("The ID (GUID) of the dashboard. Required for: get."),
    },
    async ({ action, project, team, dashboardId }) => {
      try {
        const connection = await connectionProvider();

        let resolvedProject = project;
        if (!resolvedProject) {
          const result = await elicitProject(server, connection, `Select the Azure DevOps project for ${action}.`);
          if ("response" in result) return result.response;
          resolvedProject = result.resolved;
        }

        let resolvedTeam = team;
        if (!resolvedTeam) {
          const result = await elicitTeam(server, connection, resolvedProject, `Select the Azure DevOps team for ${action}.`);
          if ("response" in result) return result.response;
          resolvedTeam = result.resolved;
        }

        const dashboardApi = await connection.getDashboardApi();
        const teamContext = { project: resolvedProject, team: resolvedTeam };

        if (action === "list") {
          const dashboards = await dashboardApi.getDashboardsByProject(teamContext);
          return { content: [{ type: "text", text: JSON.stringify(dashboards, null, 2) }] };
        }

        if (action === "get") {
          if (!dashboardId) return { content: [{ type: "text", text: "dashboardId is required for get" }], isError: true };
          const dashboard = await dashboardApi.getDashboard(teamContext, dashboardId);
          return { content: [{ type: "text", text: JSON.stringify(dashboard, null, 2) }] };
        }

        return { content: [{ type: "text", text: `Unknown action: ${action}` }], isError: true };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        const msgs: Record<string, string> = {
          list: `Error listing dashboards: ${errorMessage}`,
          get: `Error getting dashboard: ${errorMessage}`,
        };
        return { content: [{ type: "text", text: msgs[action] ?? `Error: ${errorMessage}` }], isError: true };
      }
    }
  );

  // ─── dashboard_write ────────────────────────────────────────────
  server.tool(
    DASHBOARD_TOOLS.dashboard_write,
    "Write operations for Azure DevOps team dashboards. Use the action parameter to specify the operation. Requires PAT authentication with the 'Dashboards: Read and manage' scope — this domain does not work with interactive OAuth (confirmed: TF400813 on every call).",
    {
      action: z
        .enum(["create", "update_widget"])
        .describe(
          "The action to perform. Options: create (create a new dashboard, optionally copying widgets from an existing one), update_widget (partially update one widget's settings on an existing dashboard)."
        ),
      project: z.string().optional().describe("The name or ID of the Azure DevOps project. Reuse from prior context if already known. If not provided, a project selection prompt will be shown."),
      team: z.string().optional().describe("The name or ID of the Azure DevOps team. Reuse from prior context if already known. If not provided, a team selection prompt will be shown."),
      name: z.string().optional().describe("Name of the new dashboard. Required for: create."),
      description: z.string().optional().describe("Description of the new dashboard. Used for: create."),
      widgets: z
        .string()
        .optional()
        .describe(
          "JSON array of Widget objects to place on the new dashboard, typically the (possibly modified) 'widgets' array read back from a prior 'get' call on a source dashboard. Used for: create. Omit to create an empty dashboard."
        ),
      dashboardId: z.string().optional().describe("The ID (GUID) of the dashboard containing the widget. Required for: update_widget."),
      widgetId: z.string().optional().describe("The ID (GUID) of the widget to update. Required for: update_widget."),
      widget: z
        .string()
        .optional()
        .describe(
          "Full JSON Widget object to write, with the desired fields changed (e.g. 'settings'). Use the widget object read back from a prior 'get' call on the dashboard as the base — the update API rejects the request if core fields (name, contributionId, size, position, eTag, etc.) are missing or empty, so pass the whole widget back rather than only the fields you're changing. Required for: update_widget."
        ),
    },
    async ({ action, project, team, name, description, widgets, dashboardId, widgetId, widget }) => {
      try {
        const connection = await connectionProvider();

        let resolvedProject = project;
        if (!resolvedProject) {
          const result = await elicitProject(server, connection, `Select the Azure DevOps project for ${action}.`);
          if ("response" in result) return result.response;
          resolvedProject = result.resolved;
        }

        let resolvedTeam = team;
        if (!resolvedTeam) {
          const result = await elicitTeam(server, connection, resolvedProject, `Select the Azure DevOps team for ${action}.`);
          if ("response" in result) return result.response;
          resolvedTeam = result.resolved;
        }

        const dashboardApi = await connection.getDashboardApi();
        const teamContext = { project: resolvedProject, team: resolvedTeam };

        if (action === "create") {
          if (!name) return { content: [{ type: "text", text: "name is required for create" }], isError: true };
          let parsedWidgets: Widget[] | undefined;
          if (widgets) {
            try {
              parsedWidgets = JSON.parse(widgets);
            } catch {
              return { content: [{ type: "text", text: "widgets must be a valid JSON array" }], isError: true };
            }
          }
          const created = await dashboardApi.createDashboard({ name, description, widgets: parsedWidgets }, teamContext);
          return { content: [{ type: "text", text: JSON.stringify(created, null, 2) }] };
        }

        if (action === "update_widget") {
          if (!dashboardId) return { content: [{ type: "text", text: "dashboardId is required for update_widget" }], isError: true };
          if (!widgetId) return { content: [{ type: "text", text: "widgetId is required for update_widget" }], isError: true };
          if (!widget) return { content: [{ type: "text", text: "widget is required for update_widget" }], isError: true };
          let parsedWidget: Widget;
          try {
            parsedWidget = JSON.parse(widget);
          } catch {
            return { content: [{ type: "text", text: "widget must be a valid JSON object" }], isError: true };
          }
          const updated = await dashboardApi.updateWidget(parsedWidget, teamContext, dashboardId, widgetId);
          return { content: [{ type: "text", text: JSON.stringify(updated, null, 2) }] };
        }

        return { content: [{ type: "text", text: `Unknown action: ${action}` }], isError: true };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Unknown error occurred";
        const msgs: Record<string, string> = {
          create: `Error creating dashboard: ${errorMessage}`,
          update_widget: `Error updating widget: ${errorMessage}`,
        };
        return { content: [{ type: "text", text: msgs[action] ?? `Error: ${errorMessage}` }], isError: true };
      }
    }
  );
}

export { DASHBOARD_TOOLS, configureDashboardTools };
