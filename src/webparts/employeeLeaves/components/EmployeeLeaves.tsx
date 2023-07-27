import * as React from 'react';
import styles from '../../CSS/Common.module.scss';
import { IEmployeeLeavesProps } from './IEmployeeLeavesProps';
import { sp } from "@pnp/sp";

import DataTable from 'react-data-table-component';

import "@pnp/sp/site-users/web";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface EmployeeLeavesInterface {
  WelcomeMessage: string; UserName: string; UserIsOwner: boolean; EmployeeList: IEmployeeList[]; tempEmployeeList: IEmployeeList[];
}
export interface IEmployeeList { Title: string; Year: number; AnnualLeave: number; CasualLeave: number; MedicalLeave: number; Unpaid: number; }
export interface IUserLeaves { Id: number; LeaveType: string; LeaveCount: string; RequestorName: string; DayType: string; EndDayType: string; FromDate: string; ToDate: string; Status: string; RequestDate: string }

export default class EmployeeLeaves extends React.Component<IEmployeeLeavesProps, EmployeeLeavesInterface> {

  public constructor(props: IEmployeeLeavesProps, state: EmployeeLeavesInterface) {
    super(props);
    this.state = {
      WelcomeMessage: '', UserName: '', UserIsOwner: false, EmployeeList: [] as IEmployeeList[], tempEmployeeList: [] as IEmployeeList[],
    };
  }
  public async componentDidMount() {
    let groups = await sp.web.currentUser.groups();
    let user = await sp.web.currentUser();
    const Year = new Date().getFullYear();
    groups.map((item) => { if (item.Title.indexOf("LMS Owner") !== -1) { this.setState({ UserIsOwner: true }) } });
    if (user.IsSiteAdmin == true) { this.setState({ UserIsOwner: true }) }
    await this.setState({ UserName: user.Title });
    await this.getUserMessage();
    await this.GetEmployeeLeaves(Year);
  }

  public async getUserMessage() {
    const date = new Date();
    const hour = date.getHours();
    let Output = '';
    if (hour < 12) { Output = 'Good Morning'; }
    if (hour >= 12 && hour < 16) { Output = 'Good Afternoon'; }
    if (hour >= 16 && hour < 24) { Output = 'Good Evening'; }
    await this.setState({ WelcomeMessage: Output });
  }

  public async GetEmployeeLeaves(CurrentYear: any) {
    const tempArray: any = [];
    const nextYear = CurrentYear + 1
    const FirstDay = new Date('01/01/' + CurrentYear).toISOString();
    const LastDay = new Date('01/01/' + nextYear).toISOString();

    const LeaveRequestData = await sp.web.lists.getByTitle("Leave Request").items.select("*", "RequestorName/ID", "RequestorName/Title").expand("RequestorName/Title").orderBy('Id', true).filter(`Status eq 'Approved' and (datetime'${FirstDay}' le FromDate and datetime'${LastDay}' ge ToDate)`).getAll();
    const EmployeeData: [] = await sp.web.lists.getByTitle("Employee List").items.select("*", "Employee/ID", "Employee/Title").expand("Employee/Title").filter(`Year eq '${CurrentYear}'`).orderBy('Employee', true).top(500).get();

    await EmployeeData.map(async (item: any) => {
      var AnnualLeaveHours = 0, CasualLeaveHours = 0, MedicalLeaveHours = 0, UnpaidLeaveHours = 0;

      LeaveRequestData.filter((i: any) => (i.RequestorName.Title === item.Employee.Title && i.Title === "Annual Leave") ? AnnualLeaveHours += i.LeaveCount : console.log(''))
      LeaveRequestData.filter((i: any) => (i.RequestorName.Title === item.Employee.Title && i.Title === "Casual Leave") ? CasualLeaveHours += i.LeaveCount : console.log(''))
      LeaveRequestData.filter((i: any) => (i.RequestorName.Title === item.Employee.Title && i.Title === "Medical Leave") ? MedicalLeaveHours += i.LeaveCount : console.log(''))
      LeaveRequestData.filter((i: any) => (i.RequestorName.Title === item.Employee.Title && i.Title === "Unpaid") ? UnpaidLeaveHours += i.LeaveCount : console.log(''))

      tempArray.push({ Title: item.Employee.Title, Year: item.Year, AnnualLeave: AnnualLeaveHours, CasualLeave: CasualLeaveHours, MedicalLeave: MedicalLeaveHours, Unpaid: UnpaidLeaveHours });
    });

    await this.setState({ EmployeeList: tempArray });
    await this.setState({ tempEmployeeList: tempArray });
  }
  public render(): React.ReactElement<IEmployeeLeavesProps> {
    const columns1 = [
      { name: 'Employee', selector: 'Title', sortable: true, },
      { name: 'Year', selector: 'Year', sortable: true, },
      { name: 'Annual Leave', selector: 'AnnualLeave', sortable: true, },
      { name: 'Casual Leave', selector: 'CasualLeave', sortable: true, },
      { name: 'Medical Leave', selector: 'MedicalLeave', sortable: true, },
      { name: 'Unpaid Leave', selector: 'Unpaid', sortable: true, },
    ];

    return (
      <div className={styles.sectionbox}>
        <div className={styles.section_container}>
          <div className={styles.leaves_table_box}>
            <div className={styles.webpart_box}>
              <div className={styles.webpart_title}>
                <span>Employee Leave Balance</span>
              </div>
              <div className={styles.webpart_content}>
                <div className={styles.custom_table + ' ' + styles.myleavehistory}>
                  <DataTable columns={columns1} data={this.state.EmployeeList} pagination />
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
