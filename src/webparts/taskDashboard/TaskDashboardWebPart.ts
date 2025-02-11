/**
 * SPFx Project Performance Dashboard
 * Uses Plotly.js (latest version) for interactive visualizations.
 */

import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ITaskDashboardProps } from './ITaskDashboardProps';
// import Plotly from 'https://cdn.jsdelivr.net/npm/plotly.js-dist-min';
import * as Plotly from 'plotly.js-dist-min';
import './TaskDashboard.module.css'; // Ensure this file exists in the project

interface IProjectDashboardState {
  projectData: any[];
}

export default class ProjectDashboard extends React.Component<ITaskDashboardProps, IProjectDashboardState> {
  constructor(props: ITaskDashboardProps) {
    super(props);
    this.state = { projectData: [] };
  }

  componentDidMount(): void {
    this.fetchProjectData();
  }

  fetchProjectData = async (): Promise<void> => {
    const { context } = this.props;
    const fileUrl: string = `${context.pageContext.web.absoluteUrl}/Shared Documents/PMO_Project_Data.json`;

    try {
      const response: SPHttpClientResponse = await context.spHttpClient.get(
        fileUrl,
        SPHttpClient.configurations.v1
      );
      const data: any[] = await response.json();
      this.setState({ projectData: data });
      this.createCharts(data);
    } catch (error) {
      console.error("Error fetching project data:", error);
    }
  };

  createCharts = (data: any[]): void => {
    if (!data || data.length === 0) return;

    // Pie Chart: Project Status Distribution
    const statusCounts: { [key: string]: number } = data.reduce((acc, project) => {
      acc[project["Project Status"]] = (acc[project["Project Status"]] || 0) + 1;
      return acc;
    }, {});
    Plotly.newPlot("statusChart", [{
      labels: Object.keys(statusCounts),
      values: Object.values(statusCounts),
      type: "pie"
    }], { title: "Project Status Distribution" });

    // Bar Chart: Budget vs. Actual Expenses
    Plotly.newPlot("budgetChart", [{
      x: data.map(p => p["Project Name"]),
      y: data.map(p => p["Total Project Budget"]),
      name: "Total Budget",
      type: "bar"
    }, {
      x: data.map(p => p["Project Name"]),
      y: data.map(p => p["Project Actual Expenses"]),
      name: "Actual Expenses",
      type: "bar"
    }], { title: "Budget vs. Actual Expenses", barmode: "group" });

    // Spline Chart: Project Progress Overview
    Plotly.newPlot("progressChart", [{
      x: data.map(p => p["Project Name"]),
      y: data.map(p => p["Project Progress"]),
      type: "scatter",
      mode: "lines+markers"
    }], { title: "Project Progress Overview" });

    // Histogram: Forecast vs. Previous Year Spending
    Plotly.newPlot("forecastChart", [{
      x: data.map(p => p["Next Years Forecast"]).filter(Boolean),
      type: "histogram",
      name: "Next Year’s Forecast"
    }, {
      x: data.map(p => p["Previous Year VOW"]).filter(Boolean),
      type: "histogram",
      name: "Previous Year Spending"
    }], { title: "Forecast vs. Previous Year Spending" });
  };

  // render(): React.ReactElement {
  //   return (
  //     <div className="dashboardContainer">
  //       <div id="statusChart" className="chart"></div>
  //       <div id="budgetChart" className="chart"></div>
  //       <div id="progressChart" className="chart"></div>
  //       <div id="forecastChart" className="chart"></div>
  //     </div>
  //   );
  // }
}
