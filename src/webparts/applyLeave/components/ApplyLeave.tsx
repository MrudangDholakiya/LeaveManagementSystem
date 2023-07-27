import * as React from 'react';
import styles from '../../CSS/Common.module.scss';
import { IApplyLeaveProps } from './IApplyLeaveProps';
import { sp } from "@pnp/sp";
import { Modal } from "office-ui-fabric-react/lib/Modal";
import * as moment from 'moment';
import SimpleReactValidator from 'simple-react-validator';

const LoaderIcon: any = require('../../Images/Loader.gif');

import "@pnp/sp/site-users/web";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface ApplyLeaveInterface {
  UserIsOwner: boolean; Loader: boolean; SuccessfullyPopup: boolean; UpcomingHolidayList: IUpcomingHolidayList[]; ApplyLeaveData: any; RequestorDetails: any;
  upcomingHoliday: IupcomingHolidays[];
}

export interface IUpcomingHolidayList { Title: string; Date: string; }
export interface IApplyLeaveData { Title: string; StartDate: string; EndDate: string; StartLeaveType: string; EndLeaveType: string; Reason: string; LeaveType: string; LeaveCount: string }
export interface IupcomingHolidays { Title: string; HolidayType: string; Date: string; }

export default class ApplyLeave extends React.Component<IApplyLeaveProps, ApplyLeaveInterface> {

  public validator;
  public constructor(props: IApplyLeaveProps, state: ApplyLeaveInterface) {
    super(props);
    this.state = {
      UserIsOwner: false, Loader: false, SuccessfullyPopup: false, UpcomingHolidayList: [] as IUpcomingHolidayList[], ApplyLeaveData: [], RequestorDetails: [], upcomingHoliday: [] as IupcomingHolidays[],
    };
    this.validator = new SimpleReactValidator({ autoForceUpdate: this });
  }

  public async componentDidMount() {
    let groups = await sp.web.currentUser.groups();
    let user = await sp.web.currentUser();
    var GetCurrentYear = new Date().getFullYear();

    var TempRequestorData = [];
    TempRequestorData.push({ Id: user.Id, Title: user.Title });
    await this.setState({ RequestorDetails: TempRequestorData });

    groups.map((item) => { if (item.Title.indexOf("LMS Owner") !== -1) { this.setState({ UserIsOwner: true }) } });
    if (user.IsSiteAdmin == true) { this.setState({ UserIsOwner: true }) }

    await this.getupcomingHolidays(GetCurrentYear);

  }

  public onInputChange = async (event: { target: { name: any; value: any; }; }) => {
    var EventName = event.target.name;
    var EventValue = event.target.value;
    var b = this.state.ApplyLeaveData;
    b[EventName] = EventValue;
    await this.setState({ ApplyLeaveData: b });

    if (EventName === 'StartLeaveType') {
      let StartDateState = this.state.ApplyLeaveData;
      StartDateState['StartDate'] = '';
      await this.setState({ ApplyLeaveData: StartDateState });
      (document.getElementById("StartDate") as HTMLInputElement).value = '';
    }

    if (EventName === 'StartLeaveType' && this.state.ApplyLeaveData['EndLeaveType']) {


      // END DATE TYPE
      let EndDateType = this.state.ApplyLeaveData;
      EndDateType['EndLeaveType'] = '';
      await this.setState({ ApplyLeaveData: EndDateType });
      (document.getElementById("EndLeaveType") as HTMLSelectElement).value = '';

      // NULL NO OF DAYS
      let NoOfDaysVal = this.state.ApplyLeaveData;
      NoOfDaysVal['NoOfDays'] = 0;
      await this.setState({ ApplyLeaveData: NoOfDaysVal });

    } if (EventName === 'StartDate' && this.state.ApplyLeaveData['EndDate']) {
      // END DATE
      let EndDateState = this.state.ApplyLeaveData;
      EndDateState['EndDate'] = '';
      await this.setState({ ApplyLeaveData: EndDateState });
      (document.getElementById("EndDate") as HTMLDataElement).value = '';

      // NULL NO OF DAYS
      let NoOfDaysVal = this.state.ApplyLeaveData;
      NoOfDaysVal['NoOfDays'] = 0;
      await this.setState({ ApplyLeaveData: NoOfDaysVal });
    }
    if (EventName === 'StartLeaveType' || EventName === 'EndLeaveType' || EventName === 'StartDate' || EventName === 'EndDate') {
      if ((this.state.ApplyLeaveData['StartLeaveType'] === 'First Half')) {

        // END DATE
        let EndDateState = this.state.ApplyLeaveData;
        EndDateState['EndDate'] = this.state.ApplyLeaveData['StartDate'];
        await this.setState({ ApplyLeaveData: EndDateState });

        // NULL NO OF DAYS
        let NoOfDaysVal = this.state.ApplyLeaveData;
        NoOfDaysVal['NoOfDays'] = 0;
        await this.setState({ ApplyLeaveData: NoOfDaysVal });


        (document.getElementById("EndDate") as HTMLInputElement).value = moment(this.state.ApplyLeaveData['EndDate']).format('YYYY-MM-DD');

        // END DATE TYPE
        let EndDateType = this.state.ApplyLeaveData;
        EndDateType['EndLeaveType'] = this.state.ApplyLeaveData['StartLeaveType'];
        await this.setState({ ApplyLeaveData: EndDateType });
        (document.getElementById("EndLeaveType") as HTMLSelectElement).value = this.state.ApplyLeaveData['EndLeaveType'];

      }
      if (this.state.ApplyLeaveData['StartLeaveType'] === 'Second Half' && this.state.ApplyLeaveData['EndLeaveType'] === 'Second Half') {
        // END DATE
        let EndDateState = this.state.ApplyLeaveData;
        EndDateState['EndDate'] = this.state.ApplyLeaveData['StartDate'];
        await this.setState({ ApplyLeaveData: EndDateState });
        (document.getElementById("EndDate") as HTMLInputElement).value = moment(this.state.ApplyLeaveData['EndDate']).format('YYYY-MM-DD');

        // NULL NO OF DAYS
        let NoOfDaysVal = this.state.ApplyLeaveData;
        NoOfDaysVal['NoOfDays'] = 0;
        await this.setState({ ApplyLeaveData: NoOfDaysVal });
      }
      if (this.state.ApplyLeaveData['StartLeaveType'] === 'Second Half' && this.state.ApplyLeaveData['EndLeaveType'] === 'Second Half') {
        // END DATE
        let EndDateState = this.state.ApplyLeaveData;
        EndDateState['EndDate'] = this.state.ApplyLeaveData['StartDate'];
        await this.setState({ ApplyLeaveData: EndDateState });
        (document.getElementById("EndDate") as HTMLInputElement).value = moment(this.state.ApplyLeaveData['EndDate']).format('YYYY-MM-DD');

        // NULL NO OF DAYS
        let NoOfDaysVal = this.state.ApplyLeaveData;
        NoOfDaysVal['NoOfDays'] = 0;
        await this.setState({ ApplyLeaveData: NoOfDaysVal });
      }

      if (this.state.ApplyLeaveData['StartLeaveType'] != '' && this.state.ApplyLeaveData['StartDate'] != '' && this.state.ApplyLeaveData['EndLeaveType'] != '' && this.state.ApplyLeaveData['EndDate'] != '') {
        //calculationfunction
        this.CalculationForLeave();

      }
    }
  }

  public CalculationForLeave = async () => {
    var startDate = this.state.ApplyLeaveData['StartDate'];
    var endDate = this.state.ApplyLeaveData['EndDate'];
    var startDateLeaveTypes = 0;
    var endDateLeaveTypes = 0;

    if (this.state.ApplyLeaveData['StartLeaveType'] === 'Full Day') { startDateLeaveTypes = 8; } else { startDateLeaveTypes = 4; }
    if (this.state.ApplyLeaveData['EndLeaveType'] === 'Full Day') { endDateLeaveTypes = 8; } else { endDateLeaveTypes = 4; }

    if (startDate && endDate) {
      var date_log = this.getDatesDiff(startDate, endDate, startDateLeaveTypes, endDateLeaveTypes);
      // HOURS COUNT
      if (date_log.length > 0) {
        var count = 0;
        date_log.map(el => { count += el['hours']; });
        if (this.state.ApplyLeaveData['StartDate'] && this.state.ApplyLeaveData['EndDate']) {
          if (this.state.ApplyLeaveData['StartLeaveType'] !== 'First Half') {
            // IF FULL DAY
            let FullDay = this.state.ApplyLeaveData;
            FullDay['NoOfDays'] = count / 8;
            await this.setState({ ApplyLeaveData: FullDay });
          } else {
            // IF FIRST HALF
            let HalfDay = this.state.ApplyLeaveData;
            HalfDay['NoOfDays'] = 0.5;
            await this.setState({ ApplyLeaveData: HalfDay });
          }
        } else {
          // INITIAL STATE 0
          let StateZero = this.state.ApplyLeaveData;
          StateZero['NoOfDays'] = 0;
          await this.setState({ ApplyLeaveData: StateZero });
        }
      }
    }
  }

  public getDatesDiff = (start_date: any, end_date: any, startDateLeaveTypes: number, endDateLeaveTypes: number, date_format = "YYYY-MM-DD") => {
    const getDateAsArray = (date: string) => { return moment(date.split(/\D+/), date_format); };
    const diff = getDateAsArray(end_date).diff(getDateAsArray(start_date), "days") + 1;
    const dates = [];
    for (let i = 0; i < diff; i++) {
      const nextDate = getDateAsArray(start_date).add(i, "day");
      const isWeekEndDay = nextDate.isoWeekday() > 5;
      if (i === 0) {
        if (!isWeekEndDay)
          dates.push({ data: nextDate.format(date_format), hours: startDateLeaveTypes });
      } else if (i === diff - 1) {
        if (!isWeekEndDay)
          dates.push({ data: nextDate.format(date_format), hours: endDateLeaveTypes });
      } else {
        if (!isWeekEndDay)
          dates.push({ data: nextDate.format(date_format), hours: 8 });
      }
    }
    return dates;
  }

  public saveLeaveReqData = async () => {
    if (this.validator.allValid()) {
      await this.setState({ Loader: true });
      var finalLeaveReqData = {
        "Title": this.state.ApplyLeaveData['LeaveType'], "LeaveCount": this.state.ApplyLeaveData['NoOfDays'],
        "RequestorNameId": this.state.RequestorDetails[0].Id, "DayType": this.state.ApplyLeaveData['StartLeaveType'], "EndDayType": this.state.ApplyLeaveData['EndLeaveType'],
        "FromDate": moment(this.state.ApplyLeaveData['StartDate']).format('YYYY-MM-DD'), "ToDate": moment(this.state.ApplyLeaveData['EndDate']).format('YYYY-MM-DD'),
        "OnlyYear": moment(this.state.ApplyLeaveData['StartDate']).format('YYYY'), Status:'Pending'
      };
      await sp.web.lists.getByTitle("Leave Request").items.add(JSON.parse(JSON.stringify(finalLeaveReqData))).then(async iar => {

        await this.setState({ Loader: false });
        await this.setState({ SuccessfullyPopup: true });
      })
    }
    else { this.validator.showMessages(); }
  }

  public async SuccessfullyPopupClose() {
    this.setState({ SuccessfullyPopup: false });
    location.reload();
  }

  public async getupcomingHolidays(Year: number) {
    var tempArray: { Title: any; HolidayType: any; Date: any; }[] = [];
    let today: string = (new Date()).toISOString();
    let next300day = new Date();
    next300day.setDate(next300day.getDate() + 300);
    let next300day1: string = next300day.toISOString();
    const MonthlyLeave: any[] = await sp.web.lists.getByTitle("Company Holiday").items.orderBy("Date", true).filter(`Year eq '${Year}' and datetime'${today}' le Date and datetime'${next300day1}' ge Date`).get();
    MonthlyLeave.map((item) => { tempArray.push({ Title: item.Title, HolidayType: item.HolidayType, Date: item.Date }); });
    await this.setState({ upcomingHoliday: tempArray });
  }




  public render(): React.ReactElement<IApplyLeaveProps> {

    return (
      <div className={styles.sectionbox}>
        <div className={styles.section_container}>

          <div className={styles.applyleavepage}>

            <div className={styles.applyleave_box}>
              <div className={styles.webpart_box}>
                <div className={styles.webpart_title}>
                  <span>Leave Request</span>
                </div>
                <div className={styles.webpart_content}>
                  <ul className={styles.applyforleave_list}>
                    <li className={styles.fullwdith}>
                      <label className={styles.afl_label}>Requestor</label>
                      {(this.state.RequestorDetails.length > 0) && (<>
                        {this.state.RequestorDetails.map((item: { Title: string | undefined; }, index: any) => {
                          return (<input type='text' className={styles.afl_textfld + ' ' + styles.afl_textfld_disabled} placeholder={item.Title} disabled />);
                        })}
                      </>)}
                    </li>
                    <li className={styles.halfwidth}>
                      <label className={styles.afl_label}>From <span>*</span></label>
                      <div>
                        <select className={styles.afl_textfld + ' ' + styles.afl_small} name="StartLeaveType" onChange={this.onInputChange}>
                          <option value="">Select Leave Type</option>
                          <option value="Full Day">Full Day</option>
                          <option value="First Half">First Half</option>
                          <option value="Second Half">Second Half</option>
                        </select>
                        <input type='date' className={styles.afl_textfld + ' ' + styles.afl_big} id="StartDate" name="StartDate" onChange={this.onInputChange} />
                      </div>
                      <span className={styles.error_message}>
                        {this.validator.message('Start Leave Type', this.state.ApplyLeaveData['StartLeaveType'], 'required')}
                      </span>
                      <span className={styles.error_message}>
                        {this.validator.message('Start Date', this.state.ApplyLeaveData['StartDate'], 'required')}
                      </span>
                    </li>
                    <li className={styles.halfwidth}>
                      <label className={styles.afl_label}>To <span>*</span></label>
                      <div>

                        {(!this.state.ApplyLeaveData['StartLeaveType'] || this.state.ApplyLeaveData['StartLeaveType'] == "Full Day" || this.state.ApplyLeaveData['StartLeaveType'] == "Second Half") && (<>
                          <select className={styles.afl_textfld + ' ' + styles.afl_small} id="EndLeaveType" name="EndLeaveType" onChange={this.onInputChange}>
                            <option value="">Select Leave Type</option>
                            {(this.state.ApplyLeaveData['StartLeaveType'] == "Full Day") && (<>
                              <option value="Full Day">Full Day</option>
                              <option value="First Half">First Half</option>
                            </>)}
                            {(this.state.ApplyLeaveData['StartLeaveType'] == "Second Half") && (<>
                              <option value="Full Day">Full Day</option>
                              <option value="First Half">First Half</option>
                              <option value="Second Half">Second Half</option>
                            </>)}
                          </select>
                        </>)}
                        {(this.state.ApplyLeaveData['StartLeaveType'] === 'First Half') && (<>
                          <select className={styles.afl_textfld + ' ' + styles.afl_small} id="EndLeaveType" name="EndLeaveType" disabled>
                            <option value="">Select Leave Type</option>
                            {(this.state.ApplyLeaveData['StartLeaveType'] == "First Half") && (<>
                              <option value="First Half" selected>First Half</option>
                            </>)}

                          </select>
                        </>)}

                        {(!this.state.ApplyLeaveData['StartLeaveType'] || !this.state.ApplyLeaveData['EndLeaveType']) && (<>
                          <input type='date' className={styles.afl_textfld + ' ' + styles.afl_big} id="EndDate" min={moment(this.state.ApplyLeaveData['StartDate']).add(1, 'd').format('YYYY-MM-DD')} name="EndDate" onChange={this.onInputChange} />
                        </>)}

                        {(this.state.ApplyLeaveData['StartLeaveType'] === 'First Half') && (<>
                          <input type='date' className={styles.afl_textfld + ' ' + styles.afl_big} id="EndDate" name="EndDate" disabled />
                        </>)}

                        {(this.state.ApplyLeaveData['StartLeaveType'] === 'Second Half' && this.state.ApplyLeaveData['EndLeaveType'] === 'Second Half') && (<>
                          <input type='date' className={styles.afl_textfld + ' ' + styles.afl_big} id="EndDate" name="EndDate" disabled />
                        </>)}

                        {(this.state.ApplyLeaveData['StartLeaveType'] === 'Second Half' && this.state.ApplyLeaveData['EndLeaveType'] === 'First Half') && (<>
                          <input type='date' className={styles.afl_textfld + ' ' + styles.afl_big} min={moment(this.state.ApplyLeaveData['StartDate']).add(1, 'd').format('YYYY-MM-DD')} id="EndDate" name="EndDate" onChange={this.onInputChange} />
                        </>)}

                        {((this.state.ApplyLeaveData['StartLeaveType'] === 'Full Day' && this.state.ApplyLeaveData['EndLeaveType'] === 'First Half')) && (<>
                          <input type='date' className={styles.afl_textfld + ' ' + styles.afl_big} min={moment(this.state.ApplyLeaveData['StartDate']).add(1, 'd').format('YYYY-MM-DD')} id="EndDate" name="EndDate" onChange={this.onInputChange} />
                        </>)}

                        {((this.state.ApplyLeaveData['StartLeaveType'] === 'Full Day' && this.state.ApplyLeaveData['EndLeaveType'] === 'Second Half')) && (<>
                          <input type='date' className={styles.afl_textfld + ' ' + styles.afl_big} min={moment(this.state.ApplyLeaveData['StartDate']).add(1, 'd').format('YYYY-MM-DD')} id="EndDate" name="EndDate" onChange={this.onInputChange} />
                        </>)}

                        {((this.state.ApplyLeaveData['StartLeaveType'] === 'Full Day' && this.state.ApplyLeaveData['EndLeaveType'] === 'Full Day')) && (<>
                          <input type='date' className={styles.afl_textfld + ' ' + styles.afl_big} min={moment(this.state.ApplyLeaveData['StartDate']).add(0, 'd').format('YYYY-MM-DD')} id="EndDate" name="EndDate" onChange={this.onInputChange} />
                        </>)}

                        {((this.state.ApplyLeaveData['StartLeaveType'] === 'Second Half' && this.state.ApplyLeaveData['EndLeaveType'] === 'Full Day')) && (<>
                          <input type='date' className={styles.afl_textfld + ' ' + styles.afl_big} min={moment(this.state.ApplyLeaveData['StartDate']).add(1, 'd').format('YYYY-MM-DD')} id="EndDate" name="EndDate" onChange={this.onInputChange} />
                        </>)}

                      </div>
                      <span className={styles.error_message}>
                        {this.validator.message('End Leave Type', this.state.ApplyLeaveData['EndLeaveType'], 'required')}
                      </span>
                      <span className={styles.error_message}>
                        {this.validator.message('End Date', this.state.ApplyLeaveData['EndDate'], 'required')}
                      </span>
                    </li>
                    <li className={styles.halfwidth}>
                      <label className={styles.afl_label}>No. Of Days</label>
                      <input type='text' className={styles.afl_textfld + ' ' + styles.afl_textfld_disabled} value={this.state.ApplyLeaveData['NoOfDays']} disabled />
                    </li>
                    <li className={styles.halfwidth}>
                      <label className={styles.afl_label}>Leave Type</label>
                      <select className={styles.afl_textfld} name="LeaveType" onChange={this.onInputChange}>
                        <option value="">Select Leave Type</option>
                        <option value="Annual Leave">Annual Leave</option>
                        <option value="Casual Leave">Casual Leave</option>
                        <option value="Medical Leave">Medical Leave</option>
                        <option value="Unpaid">Unpaid</option>
                      </select>
                      <span className={styles.error_message}>
                        {this.validator.message('Leave Type', this.state.ApplyLeaveData['LeaveType'], 'required')}
                      </span>
                    </li>
                    <li className={styles.fullwdith}>
                      <label className={styles.afl_label}>Reason <span>*</span></label>
                      <textarea className={styles.afl_textfld + ' ' + styles.afl_textarea} name="LeaveReason" onChange={this.onInputChange}></textarea>
                      <span className={styles.error_message}>
                        {this.validator.message('Reason', this.state.ApplyLeaveData['LeaveReason'], 'required')}
                      </span>
                    </li>
                    <li className={styles.fullwdith}>
                      <div className={styles.buttonrow}>
                        <span className={styles.afl_button + ' ' + styles.afl_submit} onClick={() => this.saveLeaveReqData()}>submit</span>
                        <a className={styles.afl_button + ' ' + styles.afl_clear} href='https://edaonca.sharepoint.com/SitePages/Dashboard.aspx'>clear</a>
                      </div>
                    </li>
                  </ul>
                </div>
              </div>
            </div>

            <div className={styles.calnder_list_box}>
              <div className={styles.webpart_box}>
                <div className={styles.webpart_title}>
                  <span>Upcoming Holidays</span>
                </div>
                <div className={styles.webpart_content}>
                  <ul className={styles.calnder_list}>
                    {this.state.upcomingHoliday.map((item, index) => {
                      return (
                        <li>
                          <div className={styles.cl_date_box}>
                            <b>{moment(new Date(item.Date)).format('DD')}</b>
                            <span>{moment(new Date(item.Date)).format('ddd')}</span>
                          </div>
                          <div className={styles.cl_content_box}>
                            <div className={styles.cl_title}>{item.Title}</div>
                            <ul className={styles.cl_location_time}>
                              <li>
                                <span>Type:</span>
                                <b>{item.HolidayType}</b>
                              </li>
                              <li>
                                <span>Date:</span>
                                <b>{moment(new Date(item.Date)).format('LL')}</b>
                              </li>
                            </ul>
                          </div>
                        </li>
                      );
                    })}
                  </ul>
                </div>
              </div>
            </div>


          </div>

          {(this.state.Loader == true) && (<div className={styles.loaderouterbox}><img src={LoaderIcon} /></div>)}

          <Modal isOpen={this.state.SuccessfullyPopup}>
            <div className={styles.modalPopup_header}>
              <h2>Leave applied successfully</h2>
              <div className={styles.closeBtnDiv} onClick={() => this.SuccessfullyPopupClose()}>
                <svg xmlns="http://www.w3.org/2000/svg" height="365pt" viewBox="0 0 365.71733 365" width="365pt">
                  <g fill="#f44336">
                    <path d="m356.339844 296.347656-286.613282-286.613281c-12.5-12.5-32.765624-12.5-45.246093 0l-15.105469 15.082031c-12.5 12.503906-12.5 32.769532 0 45.25l286.613281 286.613282c12.503907 12.5 32.769531 12.5 45.25 0l15.082031-15.082032c12.523438-12.480468 12.523438-32.75.019532-45.25zm0 0" />
                    <path d="m295.988281 9.734375-286.613281 286.613281c-12.5 12.5-12.5 32.769532 0 45.25l15.082031 15.082032c12.503907 12.5 32.769531 12.5 45.25 0l286.632813-286.59375c12.503906-12.5 12.503906-32.765626 0-45.246094l-15.082032-15.082032c-12.5-12.523437-32.765624-12.523437-45.269531-.023437zm0 0" />
                  </g>
                </svg>
              </div>
            </div>
            <div className={styles.modalPopup}>
              <p>Congratulations, you have applied for leave successfully.</p>
            </div>
            <div className={styles.modalPopup_footer}>
              <div className={styles.updatebtn} onClick={() => this.SuccessfullyPopupClose()}>Close</div>
            </div>
          </Modal>

        </div>
      </div>
    );
  }
}
