import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
// @ts-ignore
import Plotly from "plotly.js-dist";
import styles from "./TaskDashboardWebPart.module.scss";
const axios = require("axios");

export interface ITaskDashboardWebPartProps {
  description: string;
}

export default class TaskDashboardWebPart extends BaseClientSideWebPart<
  ITaskDashboardWebPartProps
> {
  private projectData: any[] = [];
  private tokenKey: string = "lawrence_task_token";
  private siteId: string = "stlawrenceparks.sharepoint.com,98a66bd0-a2ff-4c2b-9689-12f07bd3d278,d83bd597-b716-4290-b679-392487cea472";
  private driveId: string = "b!0GummP-iK0yWiRLwe9PSeJfVO9gWt5BCtnk5JIfOpHLtrW_mggf8RLP2xJHL9wv6";
  // private filePath: string = "Dashboard/Test.csv";
  private filePath: string = "Dashboard/PMO%20Project%20Budget%20Report.csv";
  private fileUrl = `https://graph.microsoft.com/v1.0/sites/${this.siteId}/drives/${this.driveId}/root:/${this.filePath}`;
  private tenantId: string = "945ebcd8-21be-4de8-9ca5-b23933913427";
  private clientId: string = "7fee58d8-db5d-4afd-9bc8-49654af795bc";
  private redirectUri: string = window.location.origin + window.location.pathname;
  private authority: string = `https://login.microsoftonline.com/${this.tenantId}`;
  private tokenEndpoint: string = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;

  public render(): void {
    this.domElement.innerHTML = `
    <div>
      <div>
        <h1 class="${styles.center}">Dashboard</h1>
        <div class="${styles.row}">
          <div id="statusChart" class="${styles.chart}"></div>
          <div id="budgetChart" class="${styles.chart}"></div>
        </div>
        <div class="${styles.row}">
          <div id="progressChart" class="${styles.chart}"></div>
          <div id="forecastChart" class="${styles.chart}"></div>
        </div>
      </div>

      <div class="${styles.title}">
        <h1 class="${styles.center}">Project Management Data Table</h1>

        <!-- Search Filter -->
        <div class="${styles.filterContainer}">
          <input 
            type="text" 
            id="searchInput"
            placeholder="Search projects..." 
            class="${styles.filterInput}"
          />
        </div>

        <div class="${styles.tableContainer}">
          <table class="${styles.dataTable}">
            <thead>
              <tr>
                <th>Project Name</th>
                <th>Status</th>
                <th>Budget ($)</th>
                <th>Expenses ($)</th>
                <th>Remaining Budget ($)</th>
                <th>Progress (%)</th>
              </tr>
            </thead>
            <tbody id="projectTableBody">
              <tr><td colspan="6">Loading data...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;


    this.fetchTaskInfo();

    // Add event listener for filtering
    document.getElementById("searchInput")?.addEventListener("input", () => {
      this.filterTable();
    });
  }

  private allProjects: any[] = []; // Store all projects for filtering

  private async fetchTaskInfo(): Promise<void> {
    const currentToken = sessionStorage.getItem(this.tokenKey);

    const graphService = new TaskDashboardWebPart();  // Initialize the MS Graph API service

    // Check if we already have a token saved in session storage
    if (currentToken) {
      // Use the stored token
      this.fetchProjectData(currentToken);
    } else {
      // Check if authorization code is already available in URL (after redirect)
      const authCode = graphService.getParameterByName('code');

      if (!authCode) {
        // Start the authorization flow to get access token if no code is present
        graphService.getAuthorizationCode();
      } else {
        // If we have the authorization code, get the access token and call Graph API
        graphService.getAccessToken(authCode)
          .then((accessToken: string) => {
            // Save the access token in session storage for subsequent use
            sessionStorage.setItem(this.tokenKey, accessToken);

            if (window.location.href !== window.location.origin + window.location.pathname) {
              const url = window.location.origin + window.location.pathname
              window.history.replaceState(null, '', url);
            }

            // Fetch data from Graph API using the new token
            this.fetchProjectData(accessToken);
          })
          .catch((error: any): void => {
            console.log('error', error);

          });
      }
    }
  }

  private async fetchProjectData(token: string): Promise<void> {
    console.log("Fetching project data...", this.fileUrl);

    try {
      // Get Access Token (You must have proper authentication setup)
      const response: any = await fetch(this.fileUrl, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
        },
      });

      if (!response.ok) {
        throw new Error("Failed to retrieve file metadata");
      }

      const metadata = await response.json();
      const downloadUrl = metadata["@microsoft.graph.downloadUrl"];

      const fileResponse = await axios.get(downloadUrl, { responseType: "stream" });

      const csvText = fileResponse.data; // Read the CSV file as text

      this.projectData = this.parseCSV(csvText); // Parse CSV data
      this.createCharts(); // Generate charts

    } catch (error) {
      console.error("Error fetching project data:", error);
    }
  }

  // Utility to generate a random string (code verifier)
  private generateCodeVerifier(): string {
    const array = new Uint32Array(56 / 2);
    window.crypto.getRandomValues(array);
    return Array.from(array, dec => ('0' + dec.toString(16)).substr(-2)).join('');
  }

  // Utility to create the code challenge
  private async generateCodeChallenge(codeVerifier: string): Promise<string> {
    const encoder = new TextEncoder();
    const data = encoder.encode(codeVerifier);
    const digest = await window.crypto.subtle.digest('SHA-256', data);
    return btoa(String.fromCharCode.apply(null, new Uint8Array(digest)))
      .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  // Method to get the authorization code from the URL
  public async getAuthorizationCode(): Promise<void> {
    const codeVerifier = this.generateCodeVerifier();
    const codeChallenge = await this.generateCodeChallenge(codeVerifier);

    // Save the code verifier in session storage (needed later)
    sessionStorage.setItem('code_task_verifier', codeVerifier);

    const authUrl = `${this.authority}/oauth2/v2.0/authorize?client_id=${this.clientId}&response_type=code&redirect_uri=${encodeURIComponent(this.redirectUri)}&response_mode=query&scope=Files.Read.All Sites.Read.All&code_challenge=${codeChallenge}&code_challenge_method=S256&state=12345`;

    window.location.href = authUrl; // Redirect user to the Microsoft login page
  }

  // Method to extract the authorization code from URL query string
  public getParameterByName(name: string): string | null {
    const url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    const regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)");
    const results = regex.exec(url);
    if (!results || !results[2]) return null;
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }

  // Method to get access token using the authorization code
  public async getAccessToken(authCode: string): Promise<string> {
    const codeVerifier = sessionStorage.getItem('code_task_verifier');
    if (!codeVerifier) {
      throw new Error("Code verifier is missing.");
    }

    const formData = new URLSearchParams();
    formData.append("client_id", this.clientId);
    formData.append("scope", "Files.Read.All Sites.Read.All");
    formData.append("code", authCode);
    formData.append("redirect_uri", this.redirectUri);
    formData.append("grant_type", "authorization_code");
    formData.append("code_verifier", codeVerifier);

    const response = await fetch(this.tokenEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: formData.toString()
    });

    if (!response.ok) {
      throw new Error("Failed to obtain access token");
    }

    const data = await response.json();
    return data.access_token;  // Return the access token
  }

  private parseCSV(csvText: string): any[] {
    const rows = csvText.split(/\r?\n/).filter(row => row.trim() !== ""); // Handles different OS line endings
    const headers = this.parseCSVRow(rows[0]); // Extract headers correctly

    return rows.slice(1).map(row => {
      const values = this.parseCSVRow(row); // Parse each row correctly
      return headers.reduce((obj, header, index) => {
        let value = values[index] !== undefined ? values[index].trim() : "N/A"; // Default empty cells to "N/A"
        if (value === "") value = "N/A"; // Ensure empty strings become "N/A"
        obj[header] = isNaN(Number(value)) ? value : Number(value); // Convert numbers, keep text as string
        return obj;
      }, {} as Record<string, string | number>);
    });
  }

  // Helper function to correctly parse CSV rows (handles commas inside quotes)
  private parseCSVRow(row: string): string[] {
    const match = row.match(/(".*?"|[^",]+|(?<=,)(?=,)|(?<=,)$)/g); // Handles empty cells correctly
    return match ? match.map(cell => cell.replace(/^"|"$/g, '')) : []; // Removes leading & trailing quotes
  }

  private populateTable(data: any[]): void {
    this.allProjects = data; // Store full dataset for filtering

    if (!data || data.length === 0) {
      const tableElement = document.getElementById('projectTableBody');
      if (tableElement) {
        tableElement.innerHTML = `<tr><td colspan="6">No data available</td></tr>`;
      }
      return;
    }

    this.renderTable(data);
  }

  private renderTable(data: any[]): void {
    const tableBody = data.map(project => `
    <tr>
      <td>${project["Project Name"]}</td>
      <td>${project["Project Status"]}</td>
      <td>${project["Total Project Budget"]?.toLocaleString()}</td>
      <td>${project["Project Actual Expenses"]?.toLocaleString()}</td>
      <td>${project["Project Remaining Budget"]?.toLocaleString()}</td>
      <td>${project["Project Progress"]}</td>
    </tr>
  `).join('');

    const tableElement = document.getElementById('projectTableBody');
    if (tableElement) {
      tableElement.innerHTML = tableBody;
    }
  }

  private filterTable(): void {
    const searchValue = (document.getElementById("searchInput") as HTMLInputElement)?.value?.toLowerCase();
    const filteredData = this.allProjects.filter(project =>
      project["Project Name"]?.toLowerCase().includes(searchValue) ||
      project["Project Status"]?.toLowerCase().includes(searchValue)
    );

    this.renderTable(filteredData);
  }

  private createCharts(): void {
    if (!this.projectData || this.projectData.length === 0) return;

    const data = this.projectData;

    // Pie Chart: Project Status Distribution
    const statusCounts: { [key: string]: number } = data.reduce(
      (acc, project) => {
        acc[project["Project Status"]] =
          (acc[project["Project Status"]] || 0) + 1;
        return acc;
      },
      {}
    );
    // Pie Chart: Project Status Distribution
    Plotly.newPlot(
      "statusChart",
      [
        {
          labels: Object.keys(statusCounts),
          values: Object.values(statusCounts),
          type: "pie",
        },
      ],
      {
        title: {
          text: "Project Status Distribution",
          font: { size: 16 },
          x: 0.5, // Center title
        },
      }
    );

    // Bar Chart: Budget vs. Actual Expenses
    Plotly.newPlot(
      "budgetChart",
      [
        {
          x: data.map((p) => p["Project Name"]),
          y: data.map((p) => p["Total Project Budget"]),
          name: "Total Budget",
          type: "bar",
        },
        {
          x: data.map((p) => p["Project Name"]),
          y: data.map((p) => p["Project Actual Expenses"]),
          name: "Actual Expenses",
          type: "bar",
        },
      ],
      {
        title: {
          text: "Budget vs. Actual Expenses",
          font: { size: 16 },
          x: 0.5,
        },
        barmode: "group",
      }
    );

    // Spline Chart: Project Progress Overview
    Plotly.newPlot(
      "progressChart",
      [
        {
          x: data.map((p) => p["Project Name"]),
          y: data.map((p) => p["Project Progress"]),
          type: "scatter",
          mode: "lines+markers",
        },
      ],
      {
        title: {
          text: "Project Progress Overview",
          font: { size: 16 },
          x: 0.5,
        },
      }
    );

    // Histogram: Forecast vs. Previous Year Spending
    Plotly.newPlot(
      "forecastChart",
      [
        {
          x: data.map((p) => p["Next Years Forecast"]).filter(Boolean),
          type: "histogram",
          name: "Next Year’s Forecast",
        },
        {
          x: data.map((p) => p["Previous Year VOW"]).filter(Boolean),
          type: "histogram",
          name: "Previous Year Spending",
        },
      ],
      {
        title: {
          text: "Histogram: Forecast vs. Previous Year Spending",
          font: { size: 16 },
          x: 0.5,
        },
      }
    );
    this.populateTable(data)
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Task Dashboard Settings" },
          groups: [
            {
              groupName: "General Settings",
              groupFields: [
                PropertyPaneTextField("description", {
                  label: "Web Part Description",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
