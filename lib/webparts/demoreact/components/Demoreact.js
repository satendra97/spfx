var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import pnp from 'sp-pnp-js';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { DateTimePicker, DateConvention } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { Web } from "sp-pnp-js";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
//import { CurrentUser } from '@pnp/sp/src/siteusers';
import { ChoiceGroup } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import * as $ from 'jquery';
require('bootstrap');
require('../css/test.css');
SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css");
SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");
var ipaddress, baseURL, dp;
var p, q, r, s, t, u, v, itemss;
var Demoreact = /** @class */ (function (_super) {
    __extends(Demoreact, _super);
    function Demoreact() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.state = {
            ActionCategory: [],
            associatedTasks: [],
            project: [],
            actions: [],
            phase: [],
            choose: [
                { key: 'cd', text: 'Customer Deployment' },
                { key: 'fd', text: 'Fennex Deployment (Internal Use)' },
                { key: 'rnd', text: 'R&D Work' }
            ],
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
        _this._getPeoplePickerItems = function (items) {
            _this.setState({
                SelectedUsers: items
            });
        };
        _this.handleChange1 = function (items) {
            _this.setState({
                SelectedDate1: items
            });
        };
        _this.handleChange2 = function (items) {
            _this.setState({
                SelectedDate2: items
            });
        };
        _this.handleChange3 = function (items) {
            _this.setState({
                SelectedDate3: items
            });
        };
        _this.onChangeProject = function (event, item) {
            console.log("Project Change : " + item.text + " " + (item.selected ? 'selected' : 'unselected'));
            _this.setState({
                Selectedproject: item.text,
                SelectedprojectId: item.key,
            }, function () {
                //this.getActionCategory(this.state.Selectedproject);
            });
            console.log(_this.state.Selectedproject);
        };
        _this.onChangeAction = function (event, item) {
            console.log("ActionCat Change : " + item.text + " " + (item.selected ? 'selected' : 'unselected'));
            _this.setState({
                Selectedactions: item.text,
                SelectedactionsId: item.key,
            }, function () {
                //this.getActionCategory(this.state.Selectedproject);
            });
            console.log(_this.state.Selectedactions);
        };
        _this.onChangeAssociatedTasks = function (event, item) {
            console.log("Task Change : " + item.text + " " + (item.selected ? 'selected' : 'unselected'));
            _this.setState({
                SelectedassociatedTasks: item.text,
                SelectedassociatedTasksId: item.key,
            }, function () {
                //this.getActionCategory(this.state.Selectedproject);
            });
            console.log(_this.state.SelectedassociatedTasks);
        };
        _this.onChangePhase = function (event, item) {
            console.log("Task Change : " + item.text + " " + (item.selected ? 'selected' : 'unselected'));
            _this.setState({
                Selectedphase: item.text,
                SelectedphaseId: item.key,
            }, function () {
                //this.getActionCategory(this.state.Selectedproject);
            });
            console.log(_this.state.Selectedproject);
        };
        _this.onChangechoice = function (event, item) {
            //   console.log(`Choice  Change : ${item.text} ${item ? 'selected' :'unselected'}`);
            _this.setState({
                Selectedchoose: item.text,
                SelectedchooseId: item.key,
            }, function () {
                //this.getActionCategory(this.state.Selectedproject);
            });
            console.log(_this.state.Selectedchoose);
        };
        return _this;
    }
    Demoreact.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", { id: "container" },
            React.createElement("form", { id: "frm" },
                React.createElement("div", { className: "row top-buff" },
                    React.createElement("div", { className: "col-lg-12" },
                        React.createElement("div", { className: "panel panel-primary" },
                            React.createElement("div", { className: "panel-heading" }, "Assign Ticket"),
                            React.createElement("div", { className: "panel-body" },
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "Title*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("input", { type: "text", id: "txttitle" }),
                                        React.createElement("p", null, "Specify the Title. Writing a good title helps users to tag and find items easily."))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2 " },
                                        React.createElement("label", null, "Project*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement(Dropdown, { placeholder: "Select Project", defaultSelectedKey: this.state.SelectedprojectId, options: this.state.project, id: "project", onChange: this.onChangeProject }))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" }),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("p", null, "Use this column to specify project. This selection helps to group action items corresponding across a project at one place."))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "ProjectType*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement(ChoiceGroup, { defaultSelectedKey: "cd", options: this.state.choose, id: "choose", onChange: this.onChangechoice }))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" }, "    "),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("p", null, "This column to specify the project type. Choose appropriate project type."))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "Action Title")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("input", { type: "text", id: "txtactntitle", required: true }))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" }, "    "),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("p", null, "Specify one line descrption of Action Item required. This title is used to search for connected tasks (if any). It's a good practice to speficy. A specific name helps in finding items easily."))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2 " },
                                        React.createElement("label", null, " ActionCategory* ")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement(Dropdown, { defaultSelectedKey: this.state.SelectedactionsId, options: this.state.actions, id: "actions", onChange: this.onChangeAction }))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" }),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("p", null, "Use this column to specify project. This selection helps to group action items corresponding across a project at one place."))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "Progress%")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("input", { type: "text", id: "txtprogress", placeholder: "0" }))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" }, "  "),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("p", null, "Specifiy your progress made for the action trakcer items"))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "Description")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("textarea", { name: "desc", id: "txtdesc", className: "form-control", rows: 6 }))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" }, "  "),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("p", null, "Add description here "))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "ClientDueDate*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement(DateTimePicker, { value: this.state.SelectedDate1, dateConvention: DateConvention.Date, 
                                            // timeConvention={TimeConvention.Hours24}
                                            onChange: this.handleChange1 }))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" }, "  "),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("p", null, "Use this column to specify the date when  deliverable is required to be developed, tested and rolled out"))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "TeamDueDate*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement(DateTimePicker, { value: this.state.SelectedDate2, dateConvention: DateConvention.Date, 
                                            //  timeConvention={TimeConvention.Hours24}
                                            onChange: this.handleChange2 }))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" }, "  "),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("p", null, "Use this column to specify the date when  deliverable is required to be developed, tested and rolled out"))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "Budgeted Hours*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("input", { type: "text", id: "txtbdgthrs", placeholder: "4 ", required: true }),
                                        React.createElement("p", null, "Use this column to enter budget hours for action tracker item."))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "Planned Hours*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("input", { type: "text", id: "txtplndhrs", placeholder: "", required: true }),
                                        React.createElement("p", null, "Use this column to specify planned hours for action tracker item."))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "Actual Hours*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("input", { type: "text", id: "txtactlhrs", placeholder: " ", required: true }),
                                        React.createElement("p", null, "Use this column to specify Actual Hours spent on the action tracker item."))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "Assigned To*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement(PeoplePicker, { context: this.props.context, 
                                            // titleText = "PeoplePicker"
                                            personSelectionLimit: 1, groupName: "", defaultSelectedUsers: this.state.SelectedUsers, showtooltip: true, isRequired: false, disabled: false, selectedItems: this._getPeoplePickerItems, showHiddenInUI: false, principalTypes: [PrincipalType.User], resolveDelay: 1000 }),
                                        React.createElement("p", null, "Specify the name of person who you think would be a best fit to work on the items."))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2 " },
                                        React.createElement("label", null, "Phase")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement(Dropdown, { defaultSelectedKey: this.state.SelectedphaseId, options: this.state.phase, id: "phase", onChange: this.onChangePhase }))),
                                React.createElement("br", null),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "ClosedDueDate*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement(DateTimePicker, { value: this.state.SelectedDate3, dateConvention: DateConvention.Date, 
                                            //timeConvention={TimeConvention.Hours24}
                                            onChange: this.handleChange3 }))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" }, "  "),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("p", null, "Specify the date on which ticket was closed. This has to be entered by Support Team"))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2 " },
                                        React.createElement("label", null, "Associated Tasks")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement(Dropdown, { defaultSelectedKey: this.state.SelectedassociatedTasksId, options: this.state.associatedTasks, id: "associatedTasks", onChange: this.onChangeAssociatedTasks }))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" }, "  "),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("p", null, "Connected Tasks"))),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("label", null, "Comments*")),
                                    React.createElement("div", { className: "col-lg-10" },
                                        React.createElement("textarea", { name: "desc", id: "txtcmnt", className: "form-control", rows: 4 }))),
                                React.createElement("br", null),
                                React.createElement("div", { className: "row top-buff" },
                                    React.createElement("div", { className: "col-lg-10 " }, " "),
                                    React.createElement("div", { className: "col-lg-2" },
                                        React.createElement("button", { type: "button", className: "btn btn-info", id: "btnsave", onClick: function () { return _this.Submit(); } }, "Save"),
                                        React.createElement("p", null, "  "),
                                        React.createElement("button", { type: "button", className: "btn btn-info", id: "btncancl", onClick: function () { return _this.Cancel(); } }, "Cancel"))))))))));
    };
    Demoreact.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        baseURL = this.props.context.pageContext.site.absoluteUrl;
                        return [4 /*yield*/, this.GetIPAddress()];
                    case 1:
                        _a.sent();
                        return [4 /*yield*/, this.getProjects()];
                    case 2:
                        _a.sent();
                        return [4 /*yield*/, this.getactnitems()];
                    case 3:
                        _a.sent();
                        return [4 /*yield*/, this.getphases()];
                    case 4:
                        _a.sent();
                        return [4 /*yield*/, this.getTasks()];
                    case 5:
                        _a.sent();
                        return [4 /*yield*/, this.getid()];
                    case 6:
                        _a.sent();
                        { /*  var queryParams = new UrlQueryParameterCollection(window.location.href);
                          itemID = queryParams.getValue("ItemID");
                          if ( itemID != null)
                            {
                              await this.getid(itemID);
                              this.setState({
                                StatusDisabled : false,
                      
                              });
                            }
                          */
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    Demoreact.prototype.GetIPAddress = function () {
        var call = $.ajax({
            url: "https://api.ipfy.org/?format=json",
            method: "GET",
            async: true,
            dataType: 'json',
            success: function (data) {
                console.log("IP Address :" + data.ip);
                ipaddress = data.ip;
            },
            error: function (textStatus, errorThrown) {
                console.log("IP Address fetch failed:" + textStatus + "--" + errorThrown);
            }
        }).responseJSON;
    };
    Demoreact.prototype.edit = function () { };
    Demoreact.prototype.Submit = function () {
        var project = this.state.SelectedprojectId;
        var date1 = this.state.SelectedDate1;
        var date2 = this.state.SelectedDate2;
        var date3 = this.state.SelectedDate3;
        var actions = this.state.SelectedactionsId;
        var associatedTasks = this.state.SelectedassociatedTasksId;
        var phase = this.state.SelectedphaseId;
        var userinfo = this.state.SelectedUsers[0];
        //let choose = this.state.SelectedchooseId;
        var userId;
        var userDetails = this.GetUserId((userinfo.secondaryText).toString());
        console.log(JSON.stringify(userDetails));
        userId = userDetails.d.Id;
        console.log(userinfo);
        var comment = $('#txtcmnt').val();
        var title = $('#txttitle').val();
        var actnttl = $('#txtactntitle').val();
        var prgrss = ($('#txtprogress').val()) / 100;
        var desc = $('#txtdesc').val();
        var bgthrs = $('#txtbdgthrs').val();
        var pldhrs = $('#txtplndhrs').val();
        var actlhrs = $('#txtactlhrs').val();
        var w = new Web(baseURL + "/fennexactntrc/");
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
        }).then(function (response) {
            console.log(response);
            { /* console.log(JSON.stringify(response.data.Id));
            w.lists.getByTitle("FnxActnTrck").items.getById(response.data.Id).update({
              Edit_Link: {
                __metadata : {type : "SP.FieldUrlValue"},
                Description : "Edit",
                Url: baseURL +
              }  */
            }
            // alert("Data saved");
            // location.reload();
        });
    };
    Demoreact.prototype.Cancel = function () {
        location.reload();
    };
    Demoreact.prototype.getProjects = function () {
        var _this = this;
        var prjcts = [];
        pnp.sp.web.lists.getByTitle("MD_Solutions").items.getAll().then(function (response) {
            console.log("Projects :" + JSON.stringify(response));
            prjcts.push({ key: "Select Project", text: "Select Project", isSelected: true });
            response.forEach(function (element) {
                prjcts.push({ key: element.ID, text: element.ProjectName });
            });
        }).then(function () {
            _this.setState({
                project: prjcts
            });
        }).then(function () {
            // this.getTasks();      
            return "Done";
        });
    };
    Demoreact.prototype.getactnitems = function () {
        var _this = this;
        pnp.sp.web.lists.getByTitle("MD_Solutions").items.getAll().then(function (response) {
            var actns = [];
            console.log("Actions :" + JSON.stringify(response));
            actns.push({ key: "Select Actions ", text: "Select Actions", isSelected: true });
            response.forEach(function (element) {
                actns.push({ key: element.ID, text: element.ActionTitle });
            });
            _this.setState({
                actions: actns
            });
        }).then(function () {
            //  this.getphases();
            return "Done";
        });
    };
    Demoreact.prototype.getphases = function () {
        var _this = this;
        pnp.sp.web.lists.getByTitle("MD_Solutions").items.getAll().then(function (response) {
            var phss = [];
            console.log("Phase :" + JSON.stringify(response));
            phss.push({ key: "Select Phase ", text: "Select Phase", isSelected: true });
            response.forEach(function (element) {
                phss.push({ key: element.ID, text: element.Phase });
            });
            _this.setState({
                phase: phss
            });
        }).then(function () {
            // this.getid();
            return "Done";
        });
    };
    Demoreact.prototype.getTasks = function () {
        var _this = this;
        pnp.sp.web.lists.getByTitle("EmployeeList").items.getAll().then(function (response) {
            var tsks = [];
            console.log("Tasks :" + JSON.stringify(response));
            tsks.push({ key: "Select Task", text: "Select Task", isSelected: true });
            response.forEach(function (elements) {
                tsks.push({ key: elements.ID, text: elements.Asttsks });
            });
            _this.setState({
                associatedTasks: tsks
            });
        }).then(function () {
            // this.getactnitems();
            return "Done";
        });
    };
    Demoreact.prototype.GetUserId = function (userName) {
        var call = $.ajax({
            url: baseURL + "/_api/web/siteusers/getbyloginname(@v)?@v=%27i:0%23.f|membership|" + userName + "%27",
            //   url: baseURL + "/_api/web/siteusers/siteusers/getbyloginname(@v)?@v='" + encodeURIComponent(userName) + "'",
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" },
            async: false,
            dataType: 'json'
        }).responseJSON;
        return call;
    };
    Demoreact.prototype.GetUserDetails = function (userId) {
        //userName format = i:0#.w|bidev\sp_admin
        var siteUrl = this.props.context.pageContext.web.absoluteUrl;
        console.log("Site URL : " + siteUrl + "/_api/web/siteusers/getbyloginname(@v)?@v=%27i:0%23.f|membership|" + userId + "%27");
        var call = $.ajax({
            url: siteUrl + "/_api/web/getuserbyid(" + parseInt(userId) + ")",
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" },
            async: false,
            dataType: 'json'
        }).responseJSON;
        //console.log("Call : " + JSON.stringify(call));
        return call.d;
    };
    Demoreact.prototype.getid = function () {
        var _this = this;
        var data;
        var w = new Web(baseURL + "/fennexactntrc/");
        w.lists.getByTitle("FnxActnTrck").items.select("Title", "Project", "ProjectType", "ActionTitle", "ActionCategory", "Progress", "Description", "ClientDueDate", "TeamDueDate", "BudgetedHours", "PlannedHours", "ActualHours", "AssignedTo", "Phase", "ClosedDate", "AssociatedTasks", "Comments").getById(88).get().then(function (item) {
            data = item;
            console.log(data);
            if (data.ClientDueDate == null || data.ClientDueDate == undefined) {
                //do nothing 
            }
            else {
                var dp = new Date(data.ClientDueDate);
                _this.setState({
                    SelectedDate1: dp
                });
            }
            if (data.TeamDueDate == null || data.TeamDueDate == undefined) {
                //do nothing 
            }
            else {
                var dp = new Date(data.TeamDueDate);
                _this.setState({
                    SelectedDate2: dp
                });
            }
            if (data.ClosedDate == null || data.ClosedDate == undefined) {
                //do nothing 
            }
            else {
                var dp = new Date(data.ClosedDate);
                _this.setState({
                    SelectedDate3: dp
                });
            }
            //peoplepicker
            var users = [];
            if (data.AssignedToId == null || data.AssignedToId == undefined) {
                //donothing
            }
            else {
                users.push(_this.GetUserDetails(data.AssignedToId).Email);
                _this.setState({
                    SelectedUsers: users
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
        }).then(function () {
            _this.getProjects();
        }).then(function () {
            _this.setState({ SelectedprojectId: data.ProjectId });
        }).then(function () {
            _this.getactnitems();
        }).then(function () {
            _this.setState({ SelectedactionsId: data.actnttlId });
        }).then(function () {
            _this.getTasks();
        }).then(function () {
            _this.setState({ SelectedassociatedTasksId: data.AssociatedTasksId });
        }).then(function () {
            _this.getphases();
        }).then(function () {
            _this.setState({ SelectedphaseId: data.PhaseId });
        });
    };
    return Demoreact;
}(React.Component));
export default Demoreact;
//# sourceMappingURL=Demoreact.js.map