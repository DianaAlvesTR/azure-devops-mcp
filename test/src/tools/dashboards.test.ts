// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { describe, expect, it } from "@jest/globals";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { configureDashboardTools } from "../../../src/tools/dashboards";

type ConnectionProviderMock = () => Promise<WebApi>;

interface DashboardApiMock {
  getDashboardsByProject: jest.Mock;
  getDashboard: jest.Mock;
  createDashboard: jest.Mock;
  updateWidget: jest.Mock;
}

interface MockConnection {
  getDashboardApi: jest.Mock;
  getCoreApi: jest.Mock;
}

describe("configureDashboardTools", () => {
  let server: McpServer;
  let connectionProvider: ConnectionProviderMock;
  let mockConnection: MockConnection;
  let mockDashboardApi: DashboardApiMock;

  beforeEach(() => {
    server = { tool: jest.fn(), server: { elicitInput: jest.fn() } } as unknown as McpServer;

    mockDashboardApi = {
      getDashboardsByProject: jest.fn(),
      getDashboard: jest.fn(),
      createDashboard: jest.fn(),
      updateWidget: jest.fn(),
    };

    mockConnection = {
      getDashboardApi: jest.fn().mockResolvedValue(mockDashboardApi),
      getCoreApi: jest.fn().mockResolvedValue({ getProjects: jest.fn(), getTeams: jest.fn() }),
    };

    connectionProvider = jest.fn().mockResolvedValue(mockConnection);
  });

  describe("tool registration", () => {
    it("registers both tools on the server", () => {
      configureDashboardTools(server, connectionProvider);
      expect(server.tool as jest.Mock).toHaveBeenCalledWith("dashboard", expect.any(String), expect.any(Object), expect.any(Function));
      expect(server.tool as jest.Mock).toHaveBeenCalledWith("dashboard_write", expect.any(String), expect.any(Object), expect.any(Function));
    });
  });

  describe("dashboard tool", () => {
    it("list action calls getDashboardsByProject with the resolved team context", async () => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard");
      if (!call) throw new Error("dashboard tool not registered");
      const [, , , handler] = call;

      const mockDashboards = [{ id: "d1", name: "Dash 1" }];
      mockDashboardApi.getDashboardsByProject.mockResolvedValue(mockDashboards);

      const result = await handler({ action: "list", project: "TaxTrade", team: "tr-tax Team" });

      expect(mockDashboardApi.getDashboardsByProject).toHaveBeenCalledWith({ project: "TaxTrade", team: "tr-tax Team" });
      expect(result.content[0].text).toBe(JSON.stringify(mockDashboards, null, 2));
    });

    it("get action calls getDashboard with the dashboardId and returns its widgets", async () => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard");
      if (!call) throw new Error("dashboard tool not registered");
      const [, , , handler] = call;

      const mockDashboard = { id: "d1", name: "Dash 1", widgets: [{ id: "w1", settings: "{}" }] };
      mockDashboardApi.getDashboard.mockResolvedValue(mockDashboard);

      const result = await handler({ action: "get", project: "TaxTrade", team: "tr-tax Team", dashboardId: "d1" });

      expect(mockDashboardApi.getDashboard).toHaveBeenCalledWith({ project: "TaxTrade", team: "tr-tax Team" }, "d1");
      expect(result.content[0].text).toBe(JSON.stringify(mockDashboard, null, 2));
    });

    it("get action returns an error when dashboardId is missing", async () => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard");
      if (!call) throw new Error("dashboard tool not registered");
      const [, , , handler] = call;

      const result = await handler({ action: "get", project: "TaxTrade", team: "tr-tax Team" });

      expect(mockDashboardApi.getDashboard).not.toHaveBeenCalled();
      expect(result.isError).toBe(true);
    });

    it("wraps API errors with an action-specific message", async () => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard");
      if (!call) throw new Error("dashboard tool not registered");
      const [, , , handler] = call;

      mockDashboardApi.getDashboardsByProject.mockRejectedValue(new Error("TF400813"));

      const result = await handler({ action: "list", project: "TaxTrade", team: "tr-tax Team" });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toBe("Error listing dashboards: TF400813");
    });
  });

  describe("dashboard_write tool", () => {
    it("create action parses the widgets JSON and calls createDashboard", async () => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard_write");
      if (!call) throw new Error("dashboard_write tool not registered");
      const [, , , handler] = call;

      const created = { id: "d2", name: "Dash Geral da Regressão 26.3 - Patch 1" };
      mockDashboardApi.createDashboard.mockResolvedValue(created);
      const widgets = [{ id: "w1", name: "Widget 1" }];

      const result = await handler({
        action: "create",
        project: "TaxTrade",
        team: "tr-tax Team",
        name: "Dash Geral da Regressão 26.3 - Patch 1",
        widgets: JSON.stringify(widgets),
      });

      expect(mockDashboardApi.createDashboard).toHaveBeenCalledWith({ name: "Dash Geral da Regressão 26.3 - Patch 1", description: undefined, widgets }, { project: "TaxTrade", team: "tr-tax Team" });
      expect(result.content[0].text).toBe(JSON.stringify(created, null, 2));
    });

    it("create action returns an error when name is missing", async () => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard_write");
      if (!call) throw new Error("dashboard_write tool not registered");
      const [, , , handler] = call;

      const result = await handler({ action: "create", project: "TaxTrade", team: "tr-tax Team" });

      expect(mockDashboardApi.createDashboard).not.toHaveBeenCalled();
      expect(result.isError).toBe(true);
    });

    it("create action returns an error when widgets is not valid JSON", async () => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard_write");
      if (!call) throw new Error("dashboard_write tool not registered");
      const [, , , handler] = call;

      const result = await handler({ action: "create", project: "TaxTrade", team: "tr-tax Team", name: "Dash", widgets: "{not json" });

      expect(mockDashboardApi.createDashboard).not.toHaveBeenCalled();
      expect(result.isError).toBe(true);
    });

    it("update_widget action parses the widget JSON and calls updateWidget with it", async () => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard_write");
      if (!call) throw new Error("dashboard_write tool not registered");
      const [, , , handler] = call;

      const updated = { id: "w1", settings: '{"planId":123}' };
      mockDashboardApi.updateWidget.mockResolvedValue(updated);
      const widget = {
        id: "w1",
        name: "Widget 1",
        contributionId: "ms.vss-dashboards-web.Microsoft.VisualStudioOnline.Dashboards.WitChartWidget",
        position: { row: 1, column: 1 },
        size: { rowSpan: 1, columnSpan: 1 },
        settings: '{"planId":123}',
        eTag: "3",
      };

      const result = await handler({
        action: "update_widget",
        project: "TaxTrade",
        team: "tr-tax Team",
        dashboardId: "d1",
        widgetId: "w1",
        widget: JSON.stringify(widget),
      });

      expect(mockDashboardApi.updateWidget).toHaveBeenCalledWith(widget, { project: "TaxTrade", team: "tr-tax Team" }, "d1", "w1");
      expect(result.content[0].text).toBe(JSON.stringify(updated, null, 2));
    });

    it.each(["dashboardId", "widgetId", "widget"])("update_widget action returns an error when %s is missing", async (missingField) => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard_write");
      if (!call) throw new Error("dashboard_write tool not registered");
      const [, , , handler] = call;

      const fullParams: Record<string, string> = {
        action: "update_widget",
        project: "TaxTrade",
        team: "tr-tax Team",
        dashboardId: "d1",
        widgetId: "w1",
        widget: JSON.stringify({ id: "w1" }),
      };
      const { [missingField]: omitted, ...params } = fullParams;
      void omitted;

      const result = await handler(params);

      expect(mockDashboardApi.updateWidget).not.toHaveBeenCalled();
      expect(result.isError).toBe(true);
    });

    it("update_widget action returns an error when widget is not valid JSON", async () => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard_write");
      if (!call) throw new Error("dashboard_write tool not registered");
      const [, , , handler] = call;

      const result = await handler({ action: "update_widget", project: "TaxTrade", team: "tr-tax Team", dashboardId: "d1", widgetId: "w1", widget: "{not json" });

      expect(mockDashboardApi.updateWidget).not.toHaveBeenCalled();
      expect(result.isError).toBe(true);
    });

    it("wraps API errors with an action-specific message", async () => {
      configureDashboardTools(server, connectionProvider);
      const call = (server.tool as jest.Mock).mock.calls.find(([toolName]) => toolName === "dashboard_write");
      if (!call) throw new Error("dashboard_write tool not registered");
      const [, , , handler] = call;

      mockDashboardApi.createDashboard.mockRejectedValue(new Error("boom"));

      const result = await handler({ action: "create", project: "TaxTrade", team: "tr-tax Team", name: "Dash" });

      expect(result.isError).toBe(true);
      expect(result.content[0].text).toBe("Error creating dashboard: boom");
    });
  });
});
