import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';


import styles from './PieChartTimeOffTypesWebPartWebPart.module.scss';
import * as strings from 'PieChartTimeOffTypesWebPartWebPartStrings';
import { sp } from "@pnp/sp";
import * as google from 'google';

export interface IPieChartTimeOffTypesWebPartWebPartProps {
  description: string;
}

export default class PieChartTimeOffTypesWebPartWebPart extends BaseClientSideWebPart<IPieChartTimeOffTypesWebPartWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {

      // other init code may be present

      sp.setup({
        spfxContext: this.context
      });
    });
  }

  private getApprovedTimeOffRequests(items): any[] {
    return items.filter(x => x.Status === "Approved");
  }

  private setApprovedTimeOffs(that: this) {
    sp.web.lists.getByTitle("TimeOffRequest").items.get().then((items: any[]) => {
      let myApprovedTimeOffRequests: any[] = that.getApprovedTimeOffRequests(items); // we need filtering at the client side, since the "Status" field is hidden
      // and we can't make the filtering from the Web API

      let timeOffTypeDaysMap = new Map<String, Number>();
      for (let i = 0; i < myApprovedTimeOffRequests.length; i++) {
        let currentTimeOffRequest = myApprovedTimeOffRequests[i];
        let startDate = new Date(currentTimeOffRequest.Start_x0020_Date);// for some reason Sharepoint
        // encodes the space of the Start Date using x0020,
        // but it doesn't encode the End Date

        let endDate = new Date(currentTimeOffRequest.EndDate);
        let millisecondsDifference = endDate.getTime() - startDate.getTime();

        let daysDifference = (millisecondsDifference / (1000 * 60 * 60 * 24)) + 1; // we need to add 1, because for example, if the time off is only 1 day, we will get 0 as a result

        let weekendsAmount: Number = that.getPeriodWeekendsDays(startDate, endDate);
        daysDifference -= weekendsAmount.valueOf();

        if (timeOffTypeDaysMap.has(currentTimeOffRequest.Timeofftype)) { // we need to add the days that are in the current key, so we don't lose the already contained days
          timeOffTypeDaysMap.set(currentTimeOffRequest.Timeofftype, timeOffTypeDaysMap.get(currentTimeOffRequest.Timeofftype).valueOf() + daysDifference);
        }
        else { // we just set the value
          timeOffTypeDaysMap.set(currentTimeOffRequest.Timeofftype, daysDifference);
        }

      }

      that.setDaysToHTML(timeOffTypeDaysMap);
    });
  }

  private getPeriodWeekendsDays(startDate: Date, endDate: Date): number { // Returns the amount of weekend days in the current period
    let amountOfWeekendsDays: number = 0;

    while (startDate < endDate) {
      let dayOfWeek = startDate.getDay();
      if (dayOfWeek == 6 || dayOfWeek == 0) { // It's either Saturday or Sunday
        amountOfWeekendsDays++;
      }

      startDate.setDate(startDate.getDate() + 1);
    }

    return amountOfWeekendsDays;
  };

  private setDaysToHTML(timeOffTypeDaysMap: Map<String, Number>) {
    google.charts.load("current", { packages: ["corechart"] });

    let that = this;
    google.charts.setOnLoadCallback(function () { that.drawChart(timeOffTypeDaysMap) });
  }

  private getDateData(key: String, timeOffTypeDaysMap: Map<String, Number>): Number {
    return timeOffTypeDaysMap.get(key) ? timeOffTypeDaysMap.get(key) : 0; //Prevents undefined behavior
  }

  private drawChart(timeOffTypeDaysMap: Map<String, Number>) {
    let that = this;

    let actualData = [
      ['Time offs', 'Days'],
      ["Paid time off", that.getDateData("Paid time off", timeOffTypeDaysMap)],
      ["Sick Leave", that.getDateData("Sick Leave", timeOffTypeDaysMap)],
      ["Unpaid time off", that.getDateData("Unpaid time off", timeOffTypeDaysMap)]
    ];

    var dataContainer = google.visualization.arrayToDataTable(actualData);

    var options = {
      title: 'Access Filtered time offs',
      is3D: true,
    };

    var chart = new google.visualization.PieChart(document.getElementById('piechart_3d'));
    chart.draw(dataContainer, options);
  }

  public render(): void {
    this.domElement.innerHTML = `<div id="piechart_3d" style="width: 900px; height: 500px;"></div>`;

    let that = this;
    that.setApprovedTimeOffs(that);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
