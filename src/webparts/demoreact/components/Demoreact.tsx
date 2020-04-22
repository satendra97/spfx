import * as React from 'react';
import styles from './Demoreact.module.scss';
import { IDemoreactProps } from './IDemoreactProps';
import { escape } from '@microsoft/sp-lodash-subset';
import pnp from 'sp-pnp-js';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { DateTimePicker, DateConvention } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { sp, Web } from "sp-pnp-js";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
//import { CurrentUser } from '@pnp/sp/src/siteusers';
import { ChoiceGroup, IChoiceGroupOption, ChoiceGroupBase } from 'office-ui-fabric-react/lib/ChoiceGroup';

import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import * as $ from 'jquery';
import { UserCustomAction } from 'sp-pnp-js/lib/sharepoint/usercustomactions';
import { UrlQueryParameterCollection } from '@microsoft/sp-core-library';

require('bootstrap');
require('../css/test.css');
SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css");
SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");

var ipaddress, baseURL , dp;
var p, q, r, s, t, u, v, itemss;
export default class Demoreact extends React.Component<IDemoreactProps, {}> {
  public state = {

    ActionCategory: [],
    associatedTasks: [],
    project: [],
    actions: [],
    phase: [],
    choose: [
      { key: 'cd', text: 'Customer Deployment' },
      { key: 'fd', text: 'Fennex Deployment (Internal Use)' },
      { key: 'rnd', text: 'R&D Work' }],


    Selectedproject: undefined,
    SelectedprojectId: undefined,

    Selectedactions: undefined,
    SelectedactionsId: undefined,

    SelectedassociatedTasks: undefined,
    SelectedassociatedTasksId: undefined,

    Selectedphase: undefined,
    SelectedphaseId: undefined,

    Selectedchoose: undefined,
    SelectedchooseId: undefined,

    //SelectedUsers : undefined,
    SelectedUsers: [],
    SelectedUserId: undefined,

    SelectedDate1: undefined,
    SelectedDate2: undefined,
    SelectedDate3: undefined,



  };
  public render(): React.ReactElement<IDemoreactProps> {
    return (
      <div id="container">
        <form id="frm">

          <div className="row top-buff">
            <div className="col-lg-12">
              <div className="panel panel-primary">
                <div className="panel-heading">
                  Assign Ticket
                  </div>
                <div className="panel-body">
                  {/* Title Part starts here   */}
                  <div className="row top-buff">
                    <div className="col-lg-2">
                      <label>Title*</label>
                    </div>
                    <div className="col-lg-10">
                      <input type="text" id="txttitle"  ></input>
                      <p>Specify the Title. Writing a good title helps users to tag and find items easily.</p>
                    </div>
                  </div>
                  {/* Title Part ends here   */}


                  {/* Project Part starts here   */}
                  <div className="row top-buff">
                    <div className="col-lg-2 ">
                      <label>Project*</label>
                    </div>
                    <div className="col-lg-10">
                      <Dropdown 
                        placeholder = "Select Project"
                        defaultSelectedKey= {this.state.SelectedprojectId}
                        options={this.state.project}
                        id="project"
                        onChange={this.onChangeProject}

                      />
                    </div>
                  </div>
                  <div className="row top-buff">
                    <div className="col-lg-2"></div>
                    <div className="col-lg-10">
                      <p>Use this column to specify project. This selection helps to group action items corresponding across a project at one place.</p>
                    </div>
                  </div>

                  {/*start here Project type*/}
                  <div className="row top-buff">
                    <div className="col-lg-2">
                      <label>ProjectType*</label>
                    </div>
                    <div className="col-lg-10">

                      <ChoiceGroup
                        defaultSelectedKey="cd"
                        options={this.state.choose}
                        id="choose"
                        onChange={this.onChangechoice}
                      />
                    </div>
                  </div>
                  <div className="row top-buff">
                    <div className="col-lg-2">    </div>
                    <div className="col-lg-10">
                      <p>This column to specify the project type. Choose appropriate project type.</p>
                    </div>
                  </div>
                  {/*Ends here Project type*/}

                  {/*Start  here Action title type*/}
                  <div className="row top-buff">
                    <div className="col-lg-2">
                      <label>Action Title</label>
                    </div>
                    <div className="col-lg-10">
                      <input type="text" id="txtactntitle" required></input>
                    </div>
                  </div>
                  <div className="row top-buff">
                    <div className="col-lg-2">    </div>
                    <div className="col-lg-10">
                      <p>Specify one line descrption of Action Item required. This title is used to search for connected tasks (if any). It's a good practice to speficy. A specific name helps in finding items easily.</p>
                    </div>
                  </div>
                  {/*Ends here Action title type*/}

                  {/*action category */}
                  <div className="row top-buff">
                    <div className="col-lg-2 ">
                      <label> ActionCategory* </label>
                    </div>
                    <div className="col-lg-10">
                      <Dropdown
                        defaultSelectedKey = {this.state.SelectedactionsId}
                        options={this.state.actions}
                        id="actions"
                        onChange={this.onChangeAction}
                      />
                    </div>
                  </div>
                  <div className="row top-buff">
                    <div className="col-lg-2"></div>
                    <div className="col-lg-10">
                      <p>Use this column to specify project. This selection helps to group action items corresponding across a project at one place.</p>
                    </div>
                  </div>
                  {/*Ends here  */}

                  {/* Progress% starts  type*/}
                  <div className="row top-buff">
                    <div className="col-lg-2">
                      <label>Progress%</label>
                    </div>
                    <div className="col-lg-10">
                      <input type="text" id="txtprogress" placeholder="0"  ></input>
                    </div>
                  </div>
                  <div className="row top-buff">
                    <div className="col-lg-2">  </div>
                    <div className="col-lg-10">
                      <p>Specifiy your progress made for the action trakcer items</p>
                    </div>
                  </div>
                  {/*Progress % Ends Here*/}

                  {/*Start  Description Here */}
                  <div className="row top-buff">
                    <div className="col-lg-2">
                      <label >Description</label>
                    </div>
                    <div className="col-lg-10" >
                      <textarea name="desc" id="txtdesc" className="form-control" rows={6} />
                    </div>
                  </div>
                  <div className="row top-buff">
                    <div className="col-lg-2">  </div>
                    <div className="col-lg-10">
                      <p>Add description here </p>
                    </div>
                  </div>

                  {/*Description % Ends Here*/}

                  {/*Client due date starts here  */}

                  <div className="row top-buff" >
                    <div className="col-lg-2" >
                      <label >ClientDueDate*</label>
                    </div>
                    <div className="col-lg-10" >
                      <DateTimePicker
                     value = {this.state.SelectedDate1 }
                        dateConvention={DateConvention.Date}
                       // timeConvention={TimeConvention.Hours24}
                        onChange={this.handleChange1}
                        
                      
                      />
                    </div>
                  </div>
                  <div className="row top-buff">
                    <div className="col-lg-2">  </div>
                    <div className="col-lg-10">
                      <p>Use this column to specify the date when  deliverable is required to be developed, tested and rolled out</p>
                    </div>
                  </div>
                  {/*Client due date ends  here  */}

                  {/*Team due date starts here  */}

                  <div className="row top-buff" >
                    <div className="col-lg-2" >
                      <label >TeamDueDate*</label>
                    </div>
                    <div className="col-lg-10" >
                      <DateTimePicker
                      value = {this.state.SelectedDate2}
                        dateConvention={DateConvention.Date}
                      //  timeConvention={TimeConvention.Hours24}
                        onChange={this.handleChange2}
                        //value={this.state. SelectedDate2}
                      />
                    </div>
                  </div>
                  <div className="row top-buff">
                    <div className="col-lg-2">  </div>
                    <div className="col-lg-10">
                      <p>Use this column to specify the date when  deliverable is required to be developed, tested and rolled out</p>
                    </div>
                  </div>
                  {/*Team  due date ends  here  */}

                  {/*BudgetHours starts here  */}
                  <div className="row top-buff">
                    <div className="col-lg-2">
                      <label>Budgeted Hours*</label>
                    </div>
                    <div className="col-lg-10">
                      <input type="text" id="txtbdgthrs" placeholder="4 " required  ></input>
                      <p>Use this column to enter budget hours for action tracker item.</p>
                    </div>
                  </div>
                  {/*BUdgeted hrs ends here  */}

                  {/*Planned Hours starts here  */}
                  <div className="row top-buff">
                    <div className="col-lg-2">
                      <label>Planned Hours*</label>
                    </div>
                    <div className="col-lg-10">
                      <input type="text" id="txtplndhrs" placeholder="" required></input>
                      <p>Use this column to specify planned hours for action tracker item.</p>
                    </div>
                  </div>
                  {/*Planned Hours ends  here  */}

                  {/*Actual Hours starts  here  */}
                  <div className="row top-buff">
                    <div className="col-lg-2">
                      <label>Actual Hours*</label>
                    </div>
                    <div className="col-lg-10">
                      <input type="text" id="txtactlhrs" placeholder=" " required></input>
                      <p>Use this column to specify Actual Hours spent on the action tracker item.</p>
                    </div>
                  </div>
                  {/*Actual  Hours ends  here  */}

                  {/*Assigned to starts  here  */}
                  <div className="row top-buff">
                    <div className="col-lg-2">
                      <label>Assigned To*</label>
                    </div>
                    <div className="col-lg-10">
                      <PeoplePicker
                        context={this.props.context}
                        // titleText = "PeoplePicker"
                        personSelectionLimit={1}
                        groupName={""}
                        defaultSelectedUsers={this.state.SelectedUsers}
                        showtooltip={true}
                        isRequired={false}
                        disabled={false}
                        selectedItems={this._getPeoplePickerItems}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                      />
                      <p>Specify the name of person who you think would be a best fit to work on the items.</p>
                    </div>
                  </div>
                  {/*Assigned to ends  here  */}


                  {/*Phase starts  here  */}
                  <div className="row top-buff">
                    <div className="col-lg-2 ">
                      <label>Phase</label>
                    </div>
                    <div className="col-lg-10">
                      <Dropdown
                        defaultSelectedKey = {this.state.SelectedphaseId}
                        options={this.state.phase}
                        id="phase"
                        onChange={this.onChangePhase}
                      />
                    </div>
                  </div>
                  <br />
                  {/*Phase Ends her   here  */}

                  {/*Close date Starts Here  */}
                  <div className="row top-buff" >
                    <div className="col-lg-2" >
                      <label >ClosedDueDate*</label>
                    </div>
                    <div className="col-lg-10" >
                      <DateTimePicker
                        value = {this.state.SelectedDate3}
                        dateConvention={DateConvention.Date}
                        //timeConvention={TimeConvention.Hours24}
                        onChange={this.handleChange3}
                     

                      />
                    </div>
                  </div>
                  <div className="row top-buff">
                    <div className="col-lg-2">  </div>
                    <div className="col-lg-10">
                      <p>Specify the date on which ticket was closed. This has to be entered by Support Team</p>
                    </div>
                  </div>

                  {/*Close date Starts Here  */}

                  {/*Associated Tasks Starts Here  */}
                  <div className="row top-buff">
                    <div className="col-lg-2 ">
                      <label>Associated Tasks</label>
                    </div>
                    <div className="col-lg-10">
                      <Dropdown
                        defaultSelectedKey = {this.state.SelectedassociatedTasksId}
                        options={this.state.associatedTasks}
                        id="associatedTasks"
                        onChange={this.onChangeAssociatedTasks}
                      />
                    </div>
                  </div>
                  <div className="row top-buff">
                    <div className="col-lg-2">  </div>
                    <div className="col-lg-10">
                      <p>Connected Tasks</p>
                    </div>
                  </div>
                  {/*Associated Tasks Ends Here  */}

                  {/*Comments start here  */}
                  <div className="row top-buff">
                    <div className="col-lg-2">
                      <label>Comments*</label>
                    </div>
                    <div className="col-lg-10">
                      <textarea name="desc" id="txtcmnt" className="form-control" rows={4} />
                    </div>
                  </div>
                  {/*Comments ends  here  */}
                  <br />

                  {/*Button start here  */}
                  <div className="row top-buff">
                    <div className="col-lg-10 "> </div>
                    <div className="col-lg-2">
                      <button type="button" className="btn btn-info" id="btnsave" onClick={() => this.Submit()}>Save</button>
                      <p>  </p>
                      <button type="button" className="btn btn-info" id="btncancl" onClick={() => this.Cancel()}>Cancel</button>
                    </div>


                  </div>

                  {/*Button ends here  */}


                </div>
              </div>
            </div>
          </div>
        </form>
      </div>
    );
  }
  public async componentDidMount() {
    baseURL = this.props.context.pageContext.site.absoluteUrl;
    await this.GetIPAddress();
    await this.getProjects();
    await this.getactnitems();
    await this.getphases();
    await this.getTasks();
 await  this.getid();
  {/*  var queryParams = new UrlQueryParameterCollection(window.location.href);
    itemID = queryParams.getValue("ItemID");
    if ( itemID != null)
      {
        await this.getid(itemID);
        this.setState({
          StatusDisabled : false,

        });
      } 
    */}

    //await this.version();
  }

  private GetIPAddress(): void {
    var call = $.ajax({
      url: "https://api.ipfy.org/?format=json",
      method: "GET",
      async: true,
      dataType: 'json',
      success: (data) => {
        console.log("IP Address :" + data.ip);
        ipaddress = data.ip;
      },
      error: (textStatus, errorThrown) => {
        console.log("IP Address fetch failed:" + textStatus + "--" + errorThrown);
      }
    }).responseJSON;
  }
  public edit(): void { }

  public Submit(): void {
    let project = this.state.SelectedprojectId;
    let date1 = this.state.SelectedDate1;
    let date2 = this.state.SelectedDate2;
    let date3 = this.state.SelectedDate3;
    let actions = this.state.SelectedactionsId;
    let associatedTasks = this.state.SelectedassociatedTasksId;
    let phase = this.state.SelectedphaseId;
    let userinfo = this.state.SelectedUsers[0];
    //let choose = this.state.SelectedchooseId;
    var userId;
    var userDetails = this.GetUserId((userinfo.secondaryText).toString());
    

    console.log(JSON.stringify(userDetails));

    userId = userDetails.d.Id;

    console.log(userinfo);
    let comment = $('#txtcmnt').val();
    let title = $('#txttitle').val();
    let actnttl = $('#txtactntitle').val();
    let prgrss = ($('#txtprogress').val()) / 100;
    let desc = $('#txtdesc').val();
    let bgthrs = $('#txtbdgthrs').val();
    let pldhrs = $('#txtplndhrs').val();
    let actlhrs = $('#txtactlhrs').val();







    let w = new Web(baseURL + "/fennexactntrc/");
    w.lists.getByTitle("FnxActnTrck").items.add({
      
      Title: title,
      ProjectId: project,
      PhaseId: phase,
      ActionTitle: actnttl,
      Comment: comment,
      ClientDueDate: date1,
      TeamDueDate: date2,
      ClosedDate: date3,
      actnttlId: actions,
      AssociatedTasksId: associatedTasks,
      Progress: prgrss,
      BudgetedHours: bgthrs,
      PlannedHours: pldhrs,
      ActualHours: actlhrs,
      AssignedToId: userId,
      Projecttype: this.state.Selectedchoose,
      Description: desc,
    }).then((response) => {

      console.log(response);

      {/* console.log(JSON.stringify(response.data.Id));
      w.lists.getByTitle("FnxActnTrck").items.getById(response.data.Id).update({
        Edit_Link: {
          __metadata : {type : "SP.FieldUrlValue"},
          Description : "Edit",
          Url: baseURL + 
        }  */}
      // alert("Data saved");
      // location.reload();

    });
  }
  private Cancel(): void {
    location.reload();
  }
  private _getPeoplePickerItems = (items): void => {
    this.setState({
      SelectedUsers: items
    });
  }
  

  private handleChange1 = (items): void => {
    this.setState({
      SelectedDate1: items

    });
  }
  private handleChange2 = (items): void => {
    this.setState({
      SelectedDate2: items

    });
  }
  private handleChange3 = (items): void => {
    this.setState({
      SelectedDate3: items

    });
  }
  private getProjects(): any {
    var prjcts = [];
    pnp.sp.web.lists.getByTitle("MD_Solutions").items.getAll().then((response) => {
      
      console.log("Projects :" + JSON.stringify(response));
      prjcts.push({ key: "Select Project", text: "Select Project", isSelected: true });
      response.forEach(element => {
        prjcts.push({ key: element.ID, text: element.ProjectName });
      });         
    }).then(() => {
      this.setState({
        project: prjcts
      });   
    }).then(() => {
  
     // this.getTasks();      
      return "Done";
    });
  }
  private getactnitems(): any {
    pnp.sp.web.lists.getByTitle("MD_Solutions").items.getAll().then((response) => {
      var actns = [];
      console.log("Actions :" + JSON.stringify(response));
      actns.push({ key: "Select Actions ", text: "Select Actions", isSelected: true });
      response.forEach(element => {
        actns.push({ key: element.ID, text: element.ActionTitle });
      });
      this.setState({
        actions: actns
      });
      
    }).then(() => {
    //  this.getphases();
      
      return "Done";
    });
  }
  private getphases(): any {
    pnp.sp.web.lists.getByTitle("MD_Solutions").items.getAll().then((response) => {
      var phss = [];
      console.log("Phase :" + JSON.stringify(response));
      phss.push({ key: "Select Phase ", text: "Select Phase", isSelected: true });
      response.forEach(element => {
        phss.push({ key: element.ID, text: element.Phase });
      });
      this.setState({
        phase: phss
      });
      
    }).then(() => {
     // this.getid();
      return "Done";
    });
  }
  private getTasks(): any {

    pnp.sp.web.lists.getByTitle("EmployeeList").items.getAll().then((response) => {
      var tsks = [];
      console.log("Tasks :" + JSON.stringify(response));
      tsks.push({ key: "Select Task", text: "Select Task", isSelected: true });
      response.forEach(elements => {
        tsks.push({ key: elements.ID, text: elements.Asttsks });
      });
      this.setState({
        associatedTasks: tsks
      });
      
    }).then(() => {
     // this.getactnitems();
      
      return "Done";
    });
  }
  private onChangeProject = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    console.log(`Project Change : ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
    this.setState({
      Selectedproject: item.text,
      SelectedprojectId: item.key,
    }, () => {
      //this.getActionCategory(this.state.Selectedproject);
    });
    console.log(this.state.Selectedproject);
  }
  private onChangeAction = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    console.log(`ActionCat Change : ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
    this.setState({
      Selectedactions: item.text,
      SelectedactionsId: item.key,
    }, () => {
      //this.getActionCategory(this.state.Selectedproject);
    });
    console.log(this.state.Selectedactions);
  }
  private onChangeAssociatedTasks = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    console.log(`Task Change : ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
    this.setState({
      SelectedassociatedTasks: item.text,
      SelectedassociatedTasksId: item.key,
    }, () => {
      //this.getActionCategory(this.state.Selectedproject);
    });
    console.log(this.state.SelectedassociatedTasks);
  }
  private onChangePhase = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    console.log(`Task Change : ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
    this.setState({
      Selectedphase: item.text,
      SelectedphaseId: item.key,
    }, () => {
      //this.getActionCategory(this.state.Selectedproject);
    });
    console.log(this.state.Selectedproject);
  }

  private onChangechoice = (event: React.FormEvent<HTMLDivElement>, item: IChoiceGroupOption): void => {
    //   console.log(`Choice  Change : ${item.text} ${item ? 'selected' :'unselected'}`);
    this.setState({
      Selectedchoose: item.text,
      SelectedchooseId: item.key,
    }, () => {
      //this.getActionCategory(this.state.Selectedproject);
    });
    console.log(this.state.Selectedchoose);
  }
  private GetUserId(userName): any {

    var call = $.ajax({

      url: baseURL + "/_api/web/siteusers/getbyloginname(@v)?@v=%27i:0%23.f|membership|" + userName + "%27",

      //   url: baseURL + "/_api/web/siteusers/siteusers/getbyloginname(@v)?@v='" + encodeURIComponent(userName) + "'",

      method: "GET",

      headers: { "Accept": "application/json; odata=verbose" },

      async: false,

      dataType: 'json'

    }).responseJSON;

    return call;

  }

    private GetUserDetails(userId) {
    //userName format = i:0#.w|bidev\sp_admin
    var siteUrl = this.props.context.pageContext.web.absoluteUrl;
    console.log("Site URL : " + siteUrl + "/_api/web/siteusers/getbyloginname(@v)?@v=%27i:0%23.f|membership|"+userId+"%27");
    var call = $.ajax({
      url: siteUrl + "/_api/web/getuserbyid(" + parseInt(userId) + ")",
      method: "GET",
      headers: { "Accept": "application/json; odata=verbose" },
      async: false,
      dataType: 'json'
    }).responseJSON;
    //console.log("Call : " + JSON.stringify(call));
    return call.d;
  }

  private getid(): any  {
     var data;
    
     let w = new Web(baseURL + "/fennexactntrc/");
      w.lists.getByTitle("FnxActnTrck").items.select("Title", "Project", "ProjectType", "ActionTitle", "ActionCategory", "Progress", "Description", "ClientDueDate", "TeamDueDate", "BudgetedHours", "PlannedHours", "ActualHours", "AssignedTo", "Phase", "ClosedDate", "AssociatedTasks", "Comments").getById(88).get().then((item: any) => {
      
      data = item;
      console.log(data); 
      if(data.ClientDueDate == null || data.ClientDueDate == undefined)
      {
        //do nothing 
      }
      else
      {
        var dp = new Date(data.ClientDueDate);

      
      this.setState(
          {
              SelectedDate1 : dp
          }
        );
      }
      if(data.TeamDueDate == null || data.TeamDueDate == undefined)
      {
      
        //do nothing 
      }
      else
      {
        var dp = new Date(data.TeamDueDate);

      
      this.setState(
          {
              SelectedDate2 : dp
          }
        );
      }
      if(data.ClosedDate == null || data.ClosedDate == undefined)
      {
        //do nothing 
      }
      
      else
      {
        var dp = new Date(data.ClosedDate);

      
      this.setState(
          {
              SelectedDate3 : dp
          }
        );
      }

      //peoplepicker
      var users  = [];
      if(data.AssignedToId == null || data.AssignedToId == undefined)
      {
        //donothing
      }
      else
      {
        users.push(this.GetUserDetails(data.AssignedToId).Email);
        this.setState({
          SelectedUsers : users 

        });
      }
      
      $('#txttitle').val(data.Title);
      $('#txtactntitle').val(data.ActionTitle);
      $('#txtprogress').val((data.Progress) * 100);
      $('#txtdesc').val(data.Description);
      $('#txtbdgthrs').val(data.BudgetedHours);
      $('#txtplndhrs').val(data.PlannedHours);
      $('#txtactlhrs').val(data.ActualHours);
      $('#txtcmnt').val(data.Comment);
     
    }).then(() =>{
        this.getProjects();
   }).then(() => {
      this.setState({SelectedprojectId : data.ProjectId});
    }).then(() => {
     this.getactnitems();
    }).then(() => {
      this.setState({SelectedactionsId : data.actnttlId});
    }).then(() =>{
      this.getTasks();
   }).then(()=>{
      this.setState({SelectedassociatedTasksId : data.AssociatedTasksId});
    }).then(() =>{
      this.getphases();
    }).then(()=>{
      this.setState({SelectedphaseId : data.PhaseId});
  
 
    });

      
}
}