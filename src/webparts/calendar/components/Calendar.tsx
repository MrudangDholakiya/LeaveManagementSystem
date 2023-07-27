import * as React from 'react';
import styles from '../../CSS/Common.module.scss';
import { ICalendarProps } from './ICalendarProps';

import * as moment from 'moment';
import tippy, { MultipleTargets, followCursor } from 'tippy.js';
import 'tippy.js/dist/tippy.css';

import { sp } from "@pnp/sp";
import "@pnp/sp/site-users/web";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/site-users/web";
import { IContextInfo } from "@pnp/sp/sites";

// import { EventInput } from '@fullcalendar/core';
import FullCalendar from "@fullcalendar/react";
import dayGridPlugin from "@fullcalendar/daygrid";
import timeGridPlugin from "@fullcalendar/timegrid";


let FilterLeaveRequestType: any = ['Annual Leave', 'Casual Leave', 'Medical Leave', 'Unpaid']

export interface ICalendar { UserIsOwner: boolean; AllLeaveRequest: ILeaveRequest[]; TempAllLeaveRequest: ILeaveRequest[]; LeaveTypeFilter: any; }
export interface ILeaveRequest { id: string; title: string; LeaveType: string; start: string; end: string; }
export interface IFilterData { AnnualLeave: boolean; CasualLeave: boolean; MedicalLeave: boolean; Unpaid: boolean; }

export default class Calendar extends React.Component<ICalendarProps, ICalendar> {

  public constructor(props: ICalendarProps) {
    super(props);
    this.state = { UserIsOwner: false, AllLeaveRequest: [] as ILeaveRequest[], TempAllLeaveRequest: [] as ILeaveRequest[], LeaveTypeFilter: [] };
  }

  public async componentDidMount() {
    let groups = await sp.web.currentUser.groups();
    let user = await sp.web.currentUser();
    groups.map((item) => { if (item.Title.indexOf("LMS Owner") !== -1) { this.setState({ UserIsOwner: true }) } });
    if (user.IsSiteAdmin == true) { this.setState({ UserIsOwner: true }) }
    let tamparray = this.state.LeaveTypeFilter;
    tamparray = {
      AnnualLeave: true,
      CasualLeave: true,
      MedicalLeave: true,
      Unpaid: true
    }

    this.setState({ LeaveTypeFilter: tamparray })
    const GetCurrentYear = new Date().getFullYear();

    
    await this.GetManagerLeaveRequests(GetCurrentYear);
  }

  public async GetManagerLeaveRequests(Year: number) {

    var tempArray: { id: any; title: any; display: any; color: any; start: any; end: any; LeaveType: any; description: any; url: any; }[] = [];
    const LeaveRequestData: any[] = await sp.web.lists.getByTitle("Leave Request").items.select("*", "RequestorName/ID", "RequestorName/Title").expand("RequestorName/Title").filter(`Status eq 'Approved' and OnlyYear eq '${Year}'`).top(5000).orderBy('Created', false).get();
    const oContext: IContextInfo = await sp.site.getContextInfo();

    LeaveRequestData.map((item) => {

      var EndDate = moment(item.ToDate).format('YYYY-MM-DD') + ('T11:59:59Z')

      if (item.Title === "Annual Leave") {
        tempArray.push({
          id: item.Id.toString(),
          title: item.RequestorName.Title + ' [' + item.Title + '] (' + item.DayType + ') to (' + item.EndDayType + ')', display: item.HRLeaveStatus,
          color: '#249103', start: new Date(item.FromDate), end: EndDate, LeaveType: item.Title, description: moment(item.FromDate).format('MM/DD/YYYY') + ' - ' + moment(item.ToDate).format('MM/DD/YYYY') + ' ' + item.RequestorName.Title  + ' [' + item.Title + '] (' + item.DayType + ') to (' + item.EndDayType + ')', url: (oContext.SiteFullUrl + '/LM/SitePages/LeaveDetails.aspx?' + item.Id.toString())
        });
      }
      if (item.Title === "Casual Leave") {
        tempArray.push({
          id: item.Id.toString(), title: item.RequestorName.Title  + ' [' + item.Title + '] (' + item.DayType + ') to (' + item.EndDayType + ')', display: item.HRLeaveStatus,
          color: '#000000', start: new Date(item.FromDate), end: EndDate, LeaveType: item.Title, description: moment(item.FromDate).format('MM/DD/YYYY') + ' - ' + moment(item.ToDate).format('MM/DD/YYYY') + ' ' + item.RequestorName.Title  + ' [' + item.Title + '] (' + item.DayType + ') to (' + item.EndDayType + ')', url: (oContext.SiteFullUrl + '/LM/SitePages/LeaveDetails.aspx?' + item.Id.toString())
        });
      }
      if (item.Title === "Medical Leave") {
        tempArray.push({
          id: item.Id.toString(), title: item.RequestorName.Title  + ' [' + item.Title + '] (' + item.DayType + ') to (' + item.EndDayType + ')', display: item.HRLeaveStatus,
          color: '#808080', start: new Date(item.FromDate), end: EndDate, LeaveType: item.Title, description: moment(item.FromDate).format('MM/DD/YYYY') + ' - ' + moment(item.ToDate).format('MM/DD/YYYY') + ' ' + item.RequestorName.Title  + ' [' + item.Title + '] (' + item.DayType + ') to (' + item.EndDayType + ')', url: (oContext.SiteFullUrl + '/LM/SitePages/LeaveDetails.aspx?' + item.Id.toString())
        });
      }
      if (item.Title === "Unpaid") {
        tempArray.push({
          id: item.Id.toString(), title: item.RequestorName.Title  + ' [' + item.Title + '] (' + item.DayType + ') to (' + item.EndDayType + ')',
          display: item.HRLeaveStatus, color: '#8a2be2', start: new Date(item.FromDate), end: EndDate, LeaveType: item.Title, description: moment(item.FromDate).format('MM/DD/YYYY') + ' - ' + moment(item.ToDate).format('MM/DD/YYYY') + ' ' + item.RequestorName.Title  + ' [' + item.Title + '] (' + item.DayType + ') to (' + item.EndDayType + ')', url: (oContext.SiteFullUrl + '/LM_UAT/SitePages/LeaveDetails.aspx?' + item.Id.toString())
        });
      }
    });
    await this.setState({ AllLeaveRequest: tempArray });
    await this.setState({ TempAllLeaveRequest: tempArray });
  }

  public _onChangeLeaveType = async (event: { target: { name: any; checked: any; }; }) => {
    let FilterAllLeaveRequest: any = this.state.TempAllLeaveRequest;
    const EventName = event.target.name;
    const EventValue = event.target.checked;

    const b = this.state.LeaveTypeFilter;
    b[EventName] = EventValue;
    await this.setState({ LeaveTypeFilter: b });

    if (EventValue) { FilterLeaveRequestType.push(EventName.replace(/([a-z])([A-Z])/g, "$1 $2")); }
    else { FilterLeaveRequestType = FilterLeaveRequestType.filter((rowId: { replaceAll: (arg0: RegExp, arg1: string) => any; }) => rowId.replaceAll(/\s/g, '') !== EventName); }
    FilterAllLeaveRequest = FilterAllLeaveRequest.filter((i:any) => FilterLeaveRequestType.indexOf(i.LeaveType) >= 0);
    await this.setState({ AllLeaveRequest: FilterAllLeaveRequest });
  }

  public handleMouseEnter = async (arg: { el: MultipleTargets; event: { extendedProps: { description: any; }; }; }) => {
    tippy(arg.el, {
      content: arg.event.extendedProps.description,
      followCursor: true,
      plugins: [followCursor],
    });
  }

  public render(): React.ReactElement<ICalendarProps> {


    return (
      <div className={styles.sectionbox}>
        <div className={styles.section_container}>

          {/* {(this.state.LeaveTypeFilter.length > 0) && ( */}
            <div className={styles.filterlisting_box}>
              <ul className={styles.filterlisting}>
                <li>
                  <label htmlFor="AnnualLeave">Annual Leave</label>
                  <input className={styles.bl_Vacation} type='checkbox' id='AnnualLeave' checked={this.state.LeaveTypeFilter.AnnualLeave} name="AnnualLeave" onChange={this._onChangeLeaveType} />
                </li>
                <li>
                  <label htmlFor="CasualLeave">Casual Leave</label>
                  <input className={styles.bl_JuryDuty} type='checkbox' id='CasualLeave' checked={this.state.LeaveTypeFilter.CasualLeave} name="CasualLeave" onChange={this._onChangeLeaveType} />
                </li>
                <li>
                  <label htmlFor="MedicalLeave">Medical Leave</label>
                  <input className={styles.bl_Bereavement} type='checkbox' id='MedicalLeave' checked={this.state.LeaveTypeFilter.MedicalLeave} name="MedicalLeave" onChange={this._onChangeLeaveType} />
                </li>
                <li>
                  <label htmlFor="Unpaid">Unpaid</label>
                  <input className={styles.bl_BusinessTravel} type='checkbox' id='Unpaid' checked={this.state.LeaveTypeFilter.Unpaid} name="Unpaid" onChange={this._onChangeLeaveType} />
                </li>
              </ul>
            </div>
          {/* )} */}

          <div className={styles.calender_box}>
            {(this.state.AllLeaveRequest.length > 0) && (<> <FullCalendar plugins={[dayGridPlugin, timeGridPlugin]} initialView="dayGridMonth" events={this.state.AllLeaveRequest} dayMaxEvents={true} headerToolbar={{ left: "prev,next", center: "title", right: "dayGridMonth,timeGridWeek,timeGridDay" }} displayEventTime={false} showNonCurrentDates={true} /> </>)}
            {(this.state.AllLeaveRequest.length === 0) && (<> <FullCalendar plugins={[dayGridPlugin, timeGridPlugin]} initialView="dayGridMonth" weekends={false} headerToolbar={{ left: "prev,next", center: "title", right: "dayGridMonth,timeGridWeek,timeGridDay" }} displayEventTime={false} showNonCurrentDates={true} /> </>)}
          </div>

        </div>
      </div>
    );
  }
}
