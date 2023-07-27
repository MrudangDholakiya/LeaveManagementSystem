import * as React from 'react';
import styles from '../../CSS/Common.module.scss';
import { IDashboardProps } from './IDashboardProps';
import { sp } from "@pnp/sp";
// import { Modal } from "office-ui-fabric-react/lib/Modal";
import * as moment from 'moment';
import DataTable from 'react-data-table-component';

import "@pnp/sp/site-users/web";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

const AddIcon: any = require('../../Images/Add.png');
const LoaderIcon: any = require('../../Images/Loader.gif');
let ClickedID: any = 0;
let ClickedStatus: any = '';

export interface ApplyLeaveInterface {
  WelcomeMessage: string; UserName: string; UserIsOwner: boolean; UserIsManager: boolean; Loader: boolean; TempUserLeaves: IUserLeaves[];
  UserLeaves: IUserLeaves[]; ApprovalLeaves: IUserLeaves[]; LeavesCount: ILeavesCount[],

  confirmUpdateStatusPopup: boolean; ApprovedSuccessFully: boolean; RejectedSuccessFully: boolean;
}
export interface IUserLeaves { Id: number; LeaveType: string; LeaveCount: string; RequestorName: string; DayType: string; EndDayType: string; FromDate: string; ToDate: string; Status: string; RequestDate: string }
export interface ILeavesCount { TotalLeave: number, Pending: number, Approved: number, Rejected: number, Cancelled: number }

export default class Dashboard extends React.Component<IDashboardProps, ApplyLeaveInterface> {

  public constructor(props: IDashboardProps, state: ApplyLeaveInterface) {
    super(props);
    this.state = {
      WelcomeMessage: '', UserName: '', UserIsOwner: false, UserIsManager: false, Loader: false, TempUserLeaves: [] as IUserLeaves[],
      UserLeaves: [] as IUserLeaves[], ApprovalLeaves: [] as IUserLeaves[], LeavesCount: [] as ILeavesCount[],
      confirmUpdateStatusPopup:false, ApprovedSuccessFully: false, RejectedSuccessFully: false
    };
  }

  public async componentDidMount() {
    let groups = await sp.web.currentUser.groups();
    let user = await sp.web.currentUser();
    var GetCurrentYear = new Date().getFullYear();

    const TempData: any[] = await sp.web.lists.getByTitle("Approver List").items.select("*", "User/ID", "User/Title").expand("User/Title").get();
    var IsManager = TempData.filter((i) => i.User.Title == user.Title);

    groups.map((item) => { if (item.Title.indexOf("LMS Owner") !== -1) { this.setState({ UserIsOwner: true }) } });
    if (user.IsSiteAdmin == true) { this.setState({ UserIsOwner: true }) }
    if (IsManager.length > 0) { this.setState({ UserIsManager: true }) }

    await this.setState({ UserName: user.Title });
    await this.getUserMessage();
    await this.getMyLeaves(GetCurrentYear, user.Title);
    await this.getAllPendingLeaves(GetCurrentYear);

  }

  public async getUserMessage() {
    const date = new Date();
    const hour = date.getHours();
    var Output = '';
    if (hour < 12) { Output = 'Good Morning'; }
    if (hour >= 12 && hour < 16) { Output = 'Good Afternoon'; }
    if (hour >= 16 && hour < 24) { Output = 'Good Evening'; }
    await this.setState({ WelcomeMessage: Output });
  }

  public async getMyLeaves(Year: number, User: string) {
    var tempArray: { Id: any; LeaveType: any; LeaveCount: any; RequestorName: any; DayType: any; EndDayType: any; FromDate: any; ToDate: any; Status: any; RequestDate: any }[] = [];
    var LeavesCount = []
    const TempData: any[] = await sp.web.lists.getByTitle("Leave Request").items.select("*", "RequestorName/ID", "RequestorName/Title").expand("RequestorName/Title").orderBy("ToDate", true).filter(`RequestorName/Title eq '${User}' and OnlyYear eq '${Year}'`).get();
    TempData.map((item) => {
      if (User == item.RequestorName.Title) {
        tempArray.push({ Id: item.Id, LeaveType: item.Title, LeaveCount: item.LeaveCount, RequestorName: item.RequestorName, DayType: item.DayType, EndDayType: item.EndDayType, FromDate: moment(item.FromDate).format('MMM DD, YYYY'), ToDate: moment(item.ToDate).format('MMM DD, YYYY'), Status: item.Status, RequestDate: moment(item.Created).format('MMM DD, YYYY') });
      }
    });

    var Pending = TempData.filter((i) => i.Status == "Pending");
    var Approved = TempData.filter((i) => i.Status == "Approved");
    var Rejected = TempData.filter((i) => i.Status == "Rejected");
    var Cancelled = TempData.filter((i) => i.Status == "Cancelled");

    LeavesCount.push({ TotalLeave: TempData.length, Pending: Pending.length, Approved: Approved.length, Rejected: Rejected.length, Cancelled: Cancelled.length })

    await this.setState({ TempUserLeaves: tempArray });
    await this.setState({ UserLeaves: tempArray });
    await this.setState({ LeavesCount: LeavesCount });

  }

  public async getAllPendingLeaves(Year: number) {
    var tempArray: { Id: any; LeaveType: any; LeaveCount: any; RequestorName: any; DayType: any; EndDayType: any; FromDate: any; ToDate: any; Status: any; RequestDate: any }[] = [];

    const TempData: any[] = await sp.web.lists.getByTitle("Leave Request").items.select("*", "RequestorName/ID", "RequestorName/Title").expand("RequestorName/Title").orderBy("ToDate", true).filter(`Status eq 'Pending' and OnlyYear eq '${Year}'`).get();
    TempData.map((item) => {
      tempArray.push({ Id: item.Id, LeaveType: item.Title, LeaveCount: item.LeaveCount, RequestorName: item.RequestorName, DayType: item.DayType, EndDayType: item.EndDayType, FromDate: moment(item.FromDate).format('MMM DD, YYYY'), ToDate: moment(item.ToDate).format('MMM DD, YYYY'), Status: item.Status, RequestDate: moment(item.Created).format('MMM DD, YYYY') });
    });
    await this.setState({ ApprovalLeaves: tempArray });

  }

  public async ShowFilterByStatus(Status: string) {
    if (Status == "All") {
      var FilterData = this.state.TempUserLeaves;
      await this.setState({ UserLeaves: FilterData })
    }
    else {
      var FilterData = this.state.TempUserLeaves.filter((i) => Status == i.Status);
      await this.setState({ UserLeaves: FilterData })
    }
  }


  public async confirmUpdateStatusPopup(status: any, ID: any) {
    ClickedID = ID;
    ClickedStatus = status;
    await this.setState({ confirmUpdateStatusPopup: true });
  }

  public async UpdateDetails() {
    if (ClickedStatus === 'Approved') {
      await sp.web.lists.getByTitle("Leave Request").items.getById(ClickedID).update({Status: ClickedStatus});
      this.setState({ ApprovedSuccessFully: true });
    } else {
      await sp.web.lists.getByTitle("Leave Request").items.getById(ClickedID).update({ Status: ClickedStatus });
      this.setState({ RejectedSuccessFully: true });
    }
    this.setState({ Loader: false });
    const GetCurrentYear = new Date().getFullYear();
    await this.getAllPendingLeaves(GetCurrentYear);
  }


  /* Start of Close Button for All Code */
  public async closePopup() {
    await this.setState({ Loader: false });
    await this.setState({ ApprovedSuccessFully: false });
    await this.setState({ RejectedSuccessFully: false });
    await this.setState({ confirmUpdateStatusPopup: false });
    ClickedID = 0;
    ClickedStatus = '';
  }
  /* End of Close Button for All Code */

  public render(): React.ReactElement<IDashboardProps> {

    const columns = [
      { name: 'Request Date', selector: 'RequestDate', sortable: true },
      { name: 'Leave Type', selector: 'LeaveType', sortable: true },
      { name: 'To Day Type', selector: 'DayType', sortable: true },
      { name: 'From Date', selector: 'FromDate', sortable: true },
      { name: 'End Day Type', selector: 'EndDayType', sortable: true },
      { name: 'To Date', selector: 'ToDate', sortable: true },
      { name: 'Leave Count', selector: 'LeaveCount', sortable: true },
      { name: 'Status', selector: 'Status', sortable: true, cell: (row: { Status: string; }) => (<div className={styles.Status + ' ' + row.Status}>{row.Status}</div>) },
    ];

    const columns1 = [
      { name: 'Request Date', selector: 'RequestDate', sortable: true },
      {
        name: 'Requested By', selector: 'LeaveType', sortable: true, cell: (row: { RequestorName: any; Id: number; }) => (
          <div className={styles.assignedtobox}>
            <div className={styles.assignedtobox_image}><img src={'/_layouts/15/userphoto.aspx?size=L&username=' + row.RequestorName.EMail} /></div>
            <div className={styles.assignedtobox_text}><span>{row.RequestorName.Title}</span><a>{row.RequestorName.EMail}</a></div>
          </div>
        )
      },
      { name: 'Leave Type', selector: 'LeaveType', sortable: true },
      { name: 'To Day Type', selector: 'DayType', sortable: true },
      { name: 'From Date', selector: 'FromDate', sortable: true },
      { name: 'End Day Type', selector: 'EndDayType', sortable: true },
      { name: 'To Date', selector: 'ToDate', sortable: true },
      { name: 'Leave Count', selector: 'LeaveCount', sortable: true },
      {
        name: 'Status', selector: 'Status', sortable: true, cell: (row: { Id: number; }) => (
          <div className={styles.buttonbox}>
            <div className={styles.approve} onClick={() => this.confirmUpdateStatusPopup('Approved', row.Id)}>Approve</div>
            <div className={styles.reject} onClick={() => this.confirmUpdateStatusPopup('Rejected', row.Id)}>Reject</div>
          </div>
        )
      },
    ];

    return (
      <div className={styles.sectionbox}>

        <div className={styles.welcome_box}>
          <div className={styles.wb_text}>
            {this.state.WelcomeMessage} <span>{this.state.UserName},</span> Welcome to <b>Leave Management System</b>
          </div>
          <ul className={styles.button_list}>
            <li>
              <a href='https://edaonca.sharepoint.com/SitePages/Apply%20Leave.aspx'>
                <img src={AddIcon} />
                <span>Apply Leave</span>
              </a>
            </li>
          </ul>
        </div>

        <div className={styles.counting_list_box}>
          {(this.state.LeavesCount.length > 0) && (
            <ul className={styles.counting_list}>
              <li onClick={() => this.ShowFilterByStatus('All')}>
                <span>All Leaves</span>
                <b>{this.state.LeavesCount[0].TotalLeave}</b>
              </li>
              <li onClick={() => this.ShowFilterByStatus('Pending')}>
                <span>Pending Leaves</span>
                <b>{this.state.LeavesCount[0].Pending}</b>
              </li>
              <li onClick={() => this.ShowFilterByStatus('Approved')}>
                <span>Approved Leaves</span>
                <b>{this.state.LeavesCount[0].Approved}</b>
              </li>
              <li onClick={() => this.ShowFilterByStatus('Rejected')}>
                <span>Rejected Leaves</span>
                <b>{this.state.LeavesCount[0].Rejected}</b>
              </li>
              <li onClick={() => this.ShowFilterByStatus('Cancelled')}>
                <span>Cancelled Leaves</span>
                <b>{this.state.LeavesCount[0].Cancelled}</b>
              </li>
            </ul>
          )}
        </div>

        <div className={styles.leaves_table_box}>
          <div className={styles.webpart_box}>
            <div className={styles.webpart_title}>
              <span>My Leave History</span>
            </div>
            <div className={styles.webpart_content}>
              <div className={styles.custom_table + ' ' + styles.myleavehistory}>
                <DataTable columns={columns} data={this.state.UserLeaves} pagination />
              </div>
            </div>
          </div>
        </div>

        {(this.state.UserIsManager == true) && (
          <div className={styles.leaves_table_box}>
            <div className={styles.webpart_box}>
              <div className={styles.webpart_title}>
                <span>My Approval</span>
              </div>
              <div className={styles.webpart_content}>
                <div className={styles.custom_table + ' ' + styles.approvalhistory}>
                  <DataTable columns={columns1} data={this.state.ApprovalLeaves} pagination />
                </div>
              </div>
            </div>
          </div>
        )}


        {(this.state.confirmUpdateStatusPopup === true) && (
          <div className={styles.custom_popup_page}>
            <div className={styles.custom_popup}>
              <div className={styles.modalPopup_header}>
                <h2>{ClickedStatus === 'Approved' ? 'Approve' : 'Reject'} Leave Request</h2>
                <div className={styles.closeBtnDiv} onClick={() => this.closePopup()}>
                  <svg xmlns="http://www.w3.org/2000/svg" height="365pt" viewBox="0 0 365.71733 365" width="365pt">
                    <g fill="#f44336">
                      <path d="m356.339844 296.347656-286.613282-286.613281c-12.5-12.5-32.765624-12.5-45.246093 0l-15.105469 15.082031c-12.5 12.503906-12.5 32.769532 0 45.25l286.613281 286.613282c12.503907 12.5 32.769531 12.5 45.25 0l15.082031-15.082032c12.523438-12.480468 12.523438-32.75.019532-45.25zm0 0" />
                      <path d="m295.988281 9.734375-286.613281 286.613281c-12.5 12.5-12.5 32.769532 0 45.25l15.082031 15.082032c12.503907 12.5 32.769531 12.5 45.25 0l286.632813-286.59375c12.503906-12.5 12.503906-32.765626 0-45.246094l-15.082032-15.082032c-12.5-12.523437-32.765624-12.523437-45.269531-.023437zm0 0" />
                    </g>
                  </svg>
                </div>
              </div>
              <div className={styles.modalPopup2}>
                {(this.state.Loader === false) && (
                  <>
                    {(this.state.ApprovedSuccessFully === false && this.state.RejectedSuccessFully === false) && (<>
                      <p>Are you sure want to {ClickedStatus === 'Approved' ? 'approve' : 'reject'} this leave request?</p>
                    </>)}
                    {(this.state.ApprovedSuccessFully === true) && (
                      <p>You have Approved Work Order Request Successfully.</p>
                    )}
                    {(this.state.RejectedSuccessFully === true) && (
                      <p>You have Rejected Work Order Request Successfully.</p>
                    )}
                  </>
                )}
                {(this.state.Loader == true) && (
                  <div className={styles.loaderbox}>
                    <img src={LoaderIcon} />
                  </div>
                )}
              </div>
              <div className={styles.modalPopup_footer}>
                {(this.state.Loader === false) && (this.state.ApprovedSuccessFully === false && this.state.RejectedSuccessFully === false) && (<div className={styles.savebtn} onClick={() => this.UpdateDetails()}>Confirm</div>)}
                <div className={styles.updatebtn} onClick={() => this.closePopup()}>Close</div>
              </div>
            </div>
          </div>
        )}

      </div>
    );
  }

}
