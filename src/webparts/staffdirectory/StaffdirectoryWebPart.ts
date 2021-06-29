import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";
import styles from "./StaffdirectoryWebPart.module.scss";
import * as strings from "StaffdirectoryWebPartStrings";
import { SPComponentLoader } from "@microsoft/sp-loader";
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js");
import * as $ from "jquery";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/profiles";
import "@pnp/sp/site-groups";
SPComponentLoader.loadScript("https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.4.min.js");

import "../../ExternalRef/CSS/style.css";

SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/select2@4.1.0-beta.1/dist/css/select2.min.css");
SPComponentLoader.loadCss("https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css");
//import "datatables";

require("datatables.net-dt");
require("datatables.net-rowgroup-dt");
SPComponentLoader.loadCss("https://cdn.datatables.net/1.10.24/css/jquery.dataTables.css");

var that: any;
declare var SP: any;
declare var SPClientPeoplePicker: any;
declare var SPClientPeoplePicker_InitStandaloneControlWrapper: any;

export interface IStaffdirectoryWebPartProps {
  description: string;
}

setTimeout(function () {
  SPComponentLoader.loadScript(
    "https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js"
  );
  SPComponentLoader.loadScript(
    "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.4/js/bootstrap.min.js"
  );
  SPComponentLoader.loadScript(
    "https://cdn.datatables.net/1.10.24/js/jquery.dataTables.js"
  );
  SPComponentLoader.loadCss(
    "https://cdn.datatables.net/rowgroup/1.0.2/css/rowGroup.dataTables.min.css"
  );
  SPComponentLoader.loadScript(
    "https://cdn.datatables.net/rowgroup/1.0.2/js/dataTables.rowGroup.min.js"
  );
}, 1000);

import "../../ExternalRef/css/alertify.min.css";
var alertify: any = require("../../ExternalRef/js/alertify.min.js");

let UserDetails = [];
let listUrl = "";
let bioAttachArr = [];
let SelectedUser = "";
let ItemID = 0;
let SelectedUserProfile = [];
let selectedUsermail = "";
let CCodeHtml = "";
let CCodeArr = [];
let OfficeAddArr = [];
let AvailEditFlag = false;
let AvailEditID = 0;
let AllAvailabilityDetails=[];
let availList=[];
var editID;
var IsgeneralStaff=false;
var IssplStaff=false;
var IsAdminStaff=false;
var currentMail="";
var ProfilePics=[];
let OfficeDetails =[]; 
var userAvailTable;
var onselecteduser="";
var onselectedusermain="";
var onselectedusersec="";
var onselecteduserTab="";
export default class StaffdirectoryWebPart extends BaseClientSideWebPart<IStaffdirectoryWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then((_) => {
      sp.setup({
        spfxContext: this.context,
      });
    });
  }

  public render(): void {
    listUrl = this.context.pageContext.web.absoluteUrl;
    currentMail = this.context.pageContext.user.email;

    onselecteduser= localStorage.getItem("EName");
    onselectedusermain= localStorage.getItem("MainSideNav");
    onselectedusersec= localStorage.getItem("secSideShow");
    onselecteduserTab= localStorage.getItem("SelectedTab");

    var siteindex = listUrl.toLocaleLowerCase().indexOf("sites");
    listUrl = listUrl.substr(siteindex - 1) + "/Lists/";
    this.domElement.innerHTML = `
    <div class="loader-section" style="display:none"> 
    <div class="loader"></div>
    </div>
    <div class="grid-section">
    <div class="left">
    <div class="left-nav">
    <div class="accordion" id="accordionExample">
    <div class="card">
      <div class="card-header nav-items SDHEmployee show" id="headingOne">
          <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseOne" aria-expanded="true" aria-controls="collapseOne"><span class="nav-icon sdh-emp"></span>SDG Employees</div>

      </div>
      <div id="collapseOne" class="clsCollapse collapse" aria-labelledby="headingOne" data-parent="#accordionExample">
        <div class="card-body">
        <div class="filter-section">
        <ul>
        <li><a href="#" class="sdhlastnamesort">By Last Name</a></li>
        <li><a href="#" class="sdhfirstnamesort">By First Name</a></li>
        <li><a href="#" class="sdhLocgrouping">By Office</a></li>
        <li><a href="#" class="sdhTitlgrouping">By Title/Staff Function</a></li>
        <li><a href="#" class="sdhAssistantgrouping">By Assistant</a></li>
        </ul>
        </div>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-header nav-items  OutsidConsultant" id="headingTwo">
          <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseTwo" aria-expanded="false" aria-controls="collapseTwo" ><span class="nav-icon out-con"></span> Outside Consultant</div>
      </div>
      <div id="collapseTwo" class="clsCollapse collapse" aria-labelledby="headingTwo" data-parent="#accordionExample">
        <div class="card-body">
        <div class="filter-section">
        <ul>
        <li><a href="#" class="OutConslastnamesort">By Last Name</a></li>
        <li><a href="#" class="OutConsFirstnamesort">By First Name</a></li>
        <li><a href="#" class="OutConsLocgrouping">By Office Affiliation</a></li>
        <li><a href="#" class="OutConsStaffgrouping">By Title/Staff Function</a></li>
        </ul>
        </div>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-header nav-items SDHAffiliates" id="headingThree">
          <div  data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseThree" aria-expanded="false" aria-controls="collapseThree"><span class="nav-icon affli"></span>Affiliates</div>
      </div>
      <div id="collapseThree" class="clsCollapse collapse" aria-labelledby="headingThree" data-parent="#accordionExample">
        <div class="card-body">
        <div class="filter-section">
        <ul>
        <li><a href="#" class="Afflastnamesort">By Last Name</a></li>
        <li><a href="#" class="AffFirstnamesort">By First Name</a></li>
        </ul>
        </div>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-header nav-items SDHAlumini" id="headingFour">
        <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseFour" aria-expanded="false" aria-controls="collapseFour"><span class="nav-icon sdh-alumini"></span>SDG Alumni</div>
      </div>
      <div id="collapseFour" class="clsCollapse collapse" aria-labelledby="headingFour" data-parent="#accordionExample">
        <div class="card-body">
        <div class="filter-section">
        <ul>
        <li><a href="#" class="SDHAlumniLastName">By Last Name</a></li>
        <li><a href="#" class="SDHAlumniFirstName">By First Name</a></li>
        <li><a href="#" class="SDHAlumniOffice">By SDG Office</a></li>
        </ul>
        </div>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-header nav-items SDHShowAll" id="headingFive">
          <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseFive" aria-expanded="false" aria-controls="collapseFive"> <span class="nav-icon show-all"></span>Show All People</div>
      </div>
      <div id="collapseFive" class="clsCollapse collapse" aria-labelledby="headingFive" data-parent="#accordionExample">
        <div class="card-body">
        <div class="filter-section">
        <ul>
        <li><a href="#" class="SDHShowAllLastName">By Last Name</a></li>
        <li><a href="#" class="SDHShowAllFirstName">By First Name</a></li>
        </ul>
        </div>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-header nav-items SDGOfficeInfo" id="headingSix">
          <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseSix" aria-expanded="false" aria-controls="collapseSix"><span class="nav-icon show-office"></span>SDG Office Info</div>
      </div>
      <!--<div id="collapseSix" class="clsCollapse collapse" aria-labelledby="headingSix" data-parent="#accordionExample">
        <div class="card-body">
        <div class="filter-section">
        <ul>
        <li><a href="#" class="SDGOfficeInfoLastName">By Last Name</a></li>
        <li><a href="#"class="SDGOfficeInfoFirstName">By First Name</a></li>
        </ul>
        </div>
        </div>
      </div>-->
    </div>
    <div class="card">
      <div class="card-header nav-items StaffAvailability" id="headingSeven">
          <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseSeven" aria-expanded="false" aria-controls="collapseSeven"><span class="nav-icon staff-avail"></span>Staff Availability</div>
      </div>
      <div id="collapseSeven" class="clsCollapse collapse" aria-labelledby="headingSeven" data-parent="#accordionExample">
        <div class="card-body">
        <div class="filter-section">
       <ul> 
        <li><a href="#" class="staffAvailOffice">By Office</a></li>
        <li><a href="#" class="staffAvailStaff">By Title/Staff Function</a></li>
        <li><a href="#" class="staffAvailAffiliation">By Staff Affiliation</a></li>
        </ul>
        </div>
        </div>
      </div>
    </div>
    <div class="card billingRateNav">
      <div class="card-header nav-items SDGBillingRate" id="headingEight">
          <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseEight" aria-expanded="false" aria-controls="collapseEight"><span class="nav-icon billing-rate"></span>Billing Rates</div>
      </div>
      <div id="collapseEight" class="clsCollapse collapse" aria-labelledby="headingEight" data-parent="#accordionExample">
        <div class="card-body">
        <div class="filter-section">
        <ul>
        <li><a href="#" class="SDGBillingRateLastName">By Last Name</a></li>
        <li><a href="#" class="SDGBillingRateFirstName">By First Name</a></li>
        <li><a href="#" class="SDGBillingRateTitle">By Title/Staff Function</a></li>
        </ul>
        </div>
        </div>
      </div>
    </div>
  </div>
    </div>
    </div>
    <div class="right">
    <div class="header-nav-back"> 
    <button class="btn-back">Back</button>
    </div>
    <div class="sdh-employee" id="SdhEmployeeDetails">
    <!-- <div class="title-section">
    <h2>Overview</h2>
    </div> -->
    <div class="title-filter-section">
    </div>
    <div class="sdh-emp-table oDataTable">

    <div class='FilterTable'>
    <div class="serchdiv">
    <label>Location:</label>
    <select id="drpLocationforEmployee">
    <option value="Select">Select</option>
    </select>
    <label>Title/Staff Function :</label>
    <select id="drpTitleforEmployee">
    <option value="Select">Select</option>
    </select>
    <label>Assistant:</label>
    <select id="drpAssistantforEmployee">
    <option value="Select">Select</option>
    </select>
    </div>
    </div>
    
    <table  id="SdhEmpTable"> 
    <thead>
    <tr>
    <th>Name</th>
    <th>First Name</th>
    <th>Last Name</th>
    <th>Phone Number</th>
    <th>Location</th>
    <th>Job Title</th>
    <th>Title/Staff Function</th>
    <th>Assistant</th>
    </tr>
    </thead>
    <tbody id="SdhEmpTbody">
    </tbody>
    </table>
    </div>
    <div class="sdh-outside-table oDataTable hide">
    <div class='FilterTable'>
    <div class="serchdiv">
    <label>Location:</label>
    <select id="drpLocationforOutside">
    <option value="Select">Select</option>
    </select>
    <label>Title/Staff Function :</label>
    <select id="drpTitleforOutside">
    <option value="Select">Select</option>
    </select>
    <label>Assistant:</label>
    <select id="drpAssistantforOutside">
    <option value="Select">Select</option>
    </select>
    </div>
    </div> 
    <table  id="SdhOutsideTable">
    <thead>
    <tr>
    <th>Name</th>
    <th>First Name</th>
    <th>Last Name</th>
    <th>Phone Number</th>
    <th>Location</th>
    <th>Job Title</th>
    <th>Title/Staff Function</th>
    <th>Assistant</th>
    </tr>
    </thead>
    <tbody id="SdhOutsideTbody">
    </tbody>
    </table>
    </div>
    <div class="sdh-Affilate-table oDataTable hide">
    <div class='FilterTable'>
    <div class="serchdiv">
    <label>Location:</label>
    <select id="drpLocationforAffiliates">
    <option value="Select">Select</option>
    </select>
    <label>Title/Staff Function :</label>
    <select id="drpTitleforAffiliates">
    <option value="Select">Select</option>
    </select>
    <label>Assistant:</label>
    <select id="drpAssistantforAffiliates">
    <option value="Select">Select</option>
    </select>
    </div>
    </div>

 
    <table  id="SdhAffilateTable">
    <thead>
    <tr>
    <th>Name</th>
    <th>First Name</th>
    <th>Last Name</th>
    <th>Phone Number</th>
    <th>Location</th>
    <th>Job Title</th>
    <th>Title/Staff Function</th>
    <th>Assistant</th>
    </tr>
    </thead>
    <tbody id="SdhAffilateTbody">
    </tbody>
    </table>
    </div>
    <div class="sdh-Allumni-table oDataTable hide">
    <div class='FilterTable'>
    <div class="serchdiv">
    <label>Location:</label>
    <select id="drpLocationforAlumni">
    <option value="Select">Select</option>
    </select>
    <label>Title/Staff Function :</label>
    <select id="drpTitleforAlumni">
    <option value="Select">Select</option>
    </select>
    <label>Assistant:</label>
    <select id="drpAssistantforAlumni">
    <option value="Select">Select</option>
    </select>
    </div>
    </div>

    
    <table  id="SdhAllumniTable">
    <thead>
    <tr>
    <th>Name</th>
    <th>First Name</th>
    <th>Last Name</th>
    <th>Phone Number</th>
    <th>Location</th>
    <th>Job Title</th>
    <th>Title/Staff Function</th>
    <th>Assistant</th>
    </tr>
    </thead>
    <tbody id="SdhAllumniTbody">
    </tbody>
    </table>
    </div>
    <div class="sdh-AllPeople-table oDataTable hide">
    <div class='FilterTable'>
    <div class="serchdiv">
    <label>Location:</label>
    <select id="drpLocationforAllPeople">
    <option value="Select">Select</option>
    </select>
    <label>Title/Staff Function :</label>
    <select id="drpTitleforAllPeople">
    <option value="Select">Select</option>
    </select>
    <label>Assistant:</label>
    <select id="drpAssistantforAllPeople">
    <option value="Select">Select</option>
    </select>
    </div>
    </div>

    
    <table  id="SdhAllPeopleTable">
    <thead>
    <tr>
    <th>Name</th>
    <th>First Name</th>
    <th>Last Name</th>
    <th>Phone Number</th>
    <th>Location</th>
    <th>Job Title</th>
    <th>Title/Staff Function</th>
    <th>Assistant</th>
    </tr>
    </thead>
    <tbody id="SdhAllPeopleTbody">
    </tbody>
    </table>
    </div>

    <div class="sdgofficeinfotable oDataTable hide">
    <table  id="SdgofficeinfoTable">
    <thead>
    <tr>
    <th>Office</th>
    <th>Phone</th>
    <th>Work Address</th>
    </tr>
    </thead>
    <tbody id="SdgofficeinfoTbody">
    </tbody>
    </table>
    </div>
    <div class="sdgbillingrateTable oDataTable hide">

    <div class='FilterTable'>
    <div class="serchdiv">
    <label>Title/Staff Function :</label>
    <select id="drpTitleforBilling">
    <option value="Select">Select</option>
    </select>
    </div>
    </div>

    <table  id="SdgBillingrateTable">
    <thead>
    <tr>
    <th>Name</th>
    <th>Title/Staff Function</th>
    <th>Daily Rate</th>
    <th>Hourly Rate</th>
    <th>Effective Date</th>
    <th>First Name</th>
    <th>Last Name</th>
    </tr>
    </thead>
    <tbody id="SdgBillingrateTbody">
    </tbody>
    </table>
    </div>
    <div class="StaffAvailabilityTable oDataTable hide">

    <div class='FilterTable'>
    <div class="serchdiv">
    <label>Title/Staff Function:</label>
    <select id="drpAffiliatesforBilling">
    <option value="Select">Select</option>
    </select>

    <label>Affiliation:</label>
    <select id="drpStaffforBilling">
    <option value="Select">Select</option>
    </select>

    </div>
    </div>

    <table id="StaffAvailabilityTable">
    <thead>
    <tr><th>User</th><th>Location</th><th>Title/Staff Function</th><th>Affiliation</th><th>Availability</th></tr>
    </thead>
    <tbody id="StaffAvailabilityTbody"></tbody>
    </table>
    </div>

    </div>

    <div class="user-profile-page hide">
    <!-- <div class="title-section">
    <h2>Employee Detail</h2>
    </div> -->
    <div class="user-profile-cover">
    <div class="cover-bg">
    <div class="profile-picture-sec">
    <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAWgAAAFoCAMAAABNO5HnAAAAvVBMVEXh4eGjo6OkpKSpqamrq6vg4ODc3Nzd3d2lpaXf39/T09PU1NTBwcHOzs7ExMS8vLysrKy+vr7R0dHFxcXX19e5ubmzs7O6urrZ2dmnp6fLy8vHx8fY2NjMzMywsLDAwMDa2trV1dWysrLIyMi0tLTCwsLKysrNzc2mpqbJycnQ0NC/v7+tra2qqqrDw8OoqKjGxsa9vb3Pz8+1tbW3t7eurq7e3t62travr6+xsbHS0tK4uLi7u7vW1tbb29sZe/uLAAAG2UlEQVR4XuzcV47dSAyG0Z+KN+ccO+ecHfe/rBl4DMNtd/cNUtXD6DtLIAhCpMiSXwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIhHnfm0cVirHTam884sVu6Q1GvPkf0heq7VE+UF5bt2y97Vat+VlRniev/EVjjp12NlgdEytLWEy5G2hepDYOt7qGob2L23Dd3valPY6dsW+jvaBOKrkm2ldBVrbag+2tYeq1oX6RxYBsF6SY3vA8to8F0roRJaZmFFK2ASWA6CiT6EhuWkoQ9gablZ6l1oW47aWoF8dpvT6FrOunoD5pa7uf6CaslyV6rqD0guzYHLRK/hwJw40Cu4MUdu9Bt8C8yR4Jt+gRbmzEKvUTicFw8kY3NonOg/aJpTTf2AWWBOBTNBkvrmWF+QNDPnZoLUNOeagpKSOVdKhK550BVa5kGLOFfMCxY92ubFuYouNC9CFdyuebKrYrsyL9hcGpgnAxVaXDJPSrGKrGreVFVkU/NmykDJj1sV2Z55s0e74hwtS9k8KvNzxY8ZozvX+L67M4/uVFwT84Kt9CPz6EjFdUqgMyCjCTSHWD4cq7jOzKMzxtGu8ddwxzzaUXHFgXkTxCqwyLyJOON0j9POc/OCpbAj+hU/Zsz9Pbk2T65VbM/mybOKbd882VexjegLPXk0L154uvF/tR5N7RjJB9bvBsLEPJgI5dCcC2P5wL3QlSClJ+bYSSpIqpljh4IkpWNzapzqB3T9vCGBuGUOtWL9hDNPizMYmjND/QIloTkSJvKB4tHRK1iaE0u9hnhgDgxi/QFJZLmLEv0FvbHlbNzTG9ApWa5KHb0J9cByFNT1DhznGOngWO9CvWQ5KdX1AXweWy7Gn/Uh9CLLQdTTCkgPLLODVCshPrSMarHWgUpkGURrl2c83drWbp+0PlRebCsvFW0G+6FtLNzXxlDuXttGrrtlbQPlacvW1ppmCDPOHgJbQ/BwpmyQnh6siHVwcJoqB3iqNx/tHY/N+pPyg7Rz83Xv0n5zuff1ppPKCSS9audf1V6i9QAAAAAAAAAAAAAAAAAAAAAAEMdyAuVeZ9I4H95/uojGgf0QjKOLT/fD88ak0ysrI6SVo9qXRWgrhIsvtaNKqs2hXNlvD0LbSDho71fKWhsxvulf2NYu+jcro42d+e0isMyCxe18R2/D6HQYWY6i4elIryE9brbMgVbzONVP2G3sBeZMsNfYFf5h715302aDIADP2Lw+CIdDQhKcGuIgKKSIk1MSMND7v6zvBvqprdqY3bWfS1itRto/O+52t+KnW+2+OdSYK+5TViS9LxxqyX07p6xUeq7hXl+WPq/AX15QI+9fDryaw5d31EP7HPGqonMb5rmvYwow/upgWTDzKYQ/C2BV3o8oSNTPYVH26FEY7zGDNfnZo0DeOYclwc6jUN4ugBVxZ0HBFp0YJoxaFK41gn7ZGxWYZtDNrSOqEK0dFLscqMbhArXuIioS3UGnHw9U5uEHFCp9quOXUGfrUSFvC11cl0p1nbK+KwHs92yFYyo2DqFEsKdq+wAqhHsqtw+hQHykescY4rnvNOC7g3TPNOEZwt3QiBuINkxpRDqEZFOaMYVgTzTkCWKFGxqyCSHVkqYsIVQQ0ZQogEwJjUkgkvNpjO8g0ZzmzCHRieacIJBLaU7qIE+bBrUhz5YGbSHPmQadIc+EBk0gT48G9SDPPQ06QZ5gQ3M2AQQa0ZwRqtCExz1kClc0ZRVCqFuacguxEhqSQC53pBlHB8HyDY3Y5BDttgnoinRoQgfinZrTuxrxgeodYiiQ+1TOz6HCy4KqLV6gREHVCqjxSsVeociaaq2hyjOVeoYyXarUhTrdZs4VeaQ6j9DIdZsXEhXpU5U+1EqoSALFtlRjC9VGHlXwRlCuTKlAWkK9rEfxehkMCB8o3EMIE1yfovUdrHiKKFb0BEMuPQrVu8CU9xNFOr3DmtcFxVm8wqBsTGHGGUxya4+CeGsHqwZjijEewDAn5Rt9dOdgWzZt6kAqMm/xylpz1EI8i3hF0SxGXQxPvJrTEHXyMuVVTF9QN+WElZuUqKPiyEodC9RV+cbKvJWos0E1TbTe4wB1l89W/GSrWY4G4G4+NUHebhwEkGGYtPgpWskQAkjSXvr8x/xlGz/RKHcr/jOrXYn/1bh0Jh7/mjfpXPALjXC+O/Av7HfzEL+nERbJZME/tpgkRYg/1Mjms48Wf1PrYzbPIIBW8aDY9j/2vsef8vz9R39bDOL/2qlDIwCBGACCOMTLl4klOpP+i4MimFe7DZy7v3rcuaYqej+f3VE1K09+AgAAAAAAAAAAAAAAAAAAAAAAgBf6wsTW1jN3CAAAAABJRU5ErkJggg==" class="profile-picture">
    </div>
    <div class="profile-name-section d-flex">
    <div class="prof-name-left">
    <p class="profile-user-name" id="UserProfileName">Sample User</p>
    <p class="profile-user-mail" id="UserProfileEmail"><span class="user-mail-icon"></span>Sample mail</p>
    </div>
    <div class="prof-name-right">
    <p class="profile-linked-id" id="UserLinkedID"><span class="user-linkedin-icon"></span>Sample mail</p>
    <p class="profile-linked-id" id="UserSkypeID"><span class="user-skype-icon"></span>Sample mail</p>
    </div>      
    </div>
    </div>
    <div class="user-details-section">
    <div class="profile-details-left">  
    <div class="user-info">
    <label>SDG Affiliation :</label>
    <div class="title-font" id="user-Designation"></div>
    </div>
    <div class="user-info">
    <label>Staff Function :</label>
    <div class="title-font" id="user-staff-function"></div>
    </div>
    <div class="user-info">
    <label>Job Title :</label>
    <div class="title-font" id="user-staff-title"></div>
    </div>
    </div>
    <div class="profile-details-right">
    <div class="user-info">
    <label id="user-phone-l">Mobile:</label>
    <div class="title-font" id="user-phone"></div>
    </div>

    <div class="user-info hide"><label>Personal Email :</label><div class="title-font" id="userpersonalmail"></div></div>
    
    </div>
    </div>
    </div>
    <div class="user-profile-tabs">
    <div class="tab-section">
    <div class="tab-header-section">
    <ul class="nav nav-tabs">
      <li id="directoryInfoTab" class="active"><a data-toggle="tab" href="#home" id="aDirectory">Directory Information</a></li>
      <li id="availabilityTab"><a data-toggle="tab" href="#menu1" id="aAvailablity">Availability</a></li>
    </ul>

    </div>
    <div>
    <div class="tab-content">
   <div id="home" class="tab-pane fade in active">
   <div class="text-right dir-edit-sec" ><button class="btn btn-edit" id="btnEdit">Edit</button></div>
     <div id="DirectoryInformation" class="d-flex view-directory">
     <div class="DInfo-left col-6">
     <div class="work-address-view">
     <h4>Office Affiliation</h4>
     <div class="d-flex"><label>Location :</label><div class="address-details lblRight" id="WLoctionDetails"></div></div>
     <div class="d-flex align-item-start"><label>Address:</label><div class="address-details lblRight" id="WAddressDetails"></div>
     </div>
     </div>

     <div class="Assistant-view" id="viewAssistant">

     </div>
     
     <div class="personal-info-view">
     <h4 class="personal-info-h">Personal Info</h4>
     <div class="address-details" id="PersonaInfo">
     <div class="d-flex"><label>Home Address :</label><div id="PAddLine" class="lblRight"></div></div>
     <div class="d-flex hide"><label>City:</label><div id="PAddCity" class="lblRight"></div></div>
     <div class="d-flex hide"><label>State:</label><div id="PAddState" class="lblRight"></div></div>
     <div class="d-flex hide"><label>Postal Code :</label><div id="PAddPCode" class="lblRight"></div></div>
     <div class="d-flex hide"><label>Country:</label><div id="PAddPCountry" class="lblRight"></div></div>
     <div><label>Nationality :</label><label id="citizenship" class="lblRight"></label></div>
     <div class="d-flex"><label>Significant Other :</label><div id="PSignOther" class="lblRight"></div></div>
     <div class="d-flex"><label>Children :</label><div id="PChildren" class="lblRight"></div></div>

     </div>
     </div>
     <div class="contact-info-view">
     <h4 class="contact-info-l">Contact Info</h4>
     <div class="address-details" id="ContactInfo">
     <div class="d-flex"><label id="homeNoviewl">Home No :</label><div id="homeNoview" class="lblRight"></div></div>
     <div class="d-flex"><label id="emergencyNoviewl">Emergency No :</label><div id="emergencyNoview" class="lblRight"></div></div>
     <div class="d-flex"><label id="officeNoviewl">Office No :</label><div id="officeNoview" class="lblRight"></div></div>
     <!--<div class="d-flex"><label>Significant Other :</label><div id="PSignOther" class="lblRight"></div></div>
     <div class="d-flex"><label>Children :</label><div id="PChildren" class="lblRight"></div></div>-->
     </div>
     </div>


     <div class="StaffStatus-view">
     <h4>Staff Status</h4>
     <p class="lblRight" id="staffStatus"></p>
     <div id="workscheduleViewSec">
     <div class="d-flex"><label>Work Schedule</label><p class="lblRight" id="workSchedule"></p></div>

     </div>
     </div>
     <div class="citizen-info">

     <div class="address-details" id="CitizenInfo">
     <!--<div><label>Nationality :</label><label id="citizenship" class="lblRight"></label></div>-->
     </div>
     </div>
     </div>
     <div class="DInfo-right col-6">
     <h4 class="BRHead" style="display:none">Billing Rate</h4>
     <div class="user-billing-rates hide">
     
     <div id="BillingRateDetails">
     <div class="billing-rates"><label>USD Daily Rate</label><div class="usd-daily-rate" id="UsdDailyRate"></div></div>
     <div class="billing-rates"><label>USD Hourly Rate</label><div class="usd-hourly-rate" id="UsdHourlyRate"></div></div>
     <div class="billing-rates"><label>EUR Daily Rate</label><div class="eur-daily-rate" id="EURDailyRate"></div></div>
     <div class="billing-rates"><label>EUR Hourly Rate</label><div class="eur-hourly-rate" id="EURHourlyRate"></div></div>
     <div class="billing-effective-date"><label>Effective Date</label><div class="effective-date" id="EffectiveDate"></div></div>
     <div class="billing-rates"><label>Comments</label><div class="w-100"><textarea class="Billing-comments" id="BillingRateComments"></textarea></div></div>
     </div>
     </div>
     <div class="Biography-Experience-view">
     <h4>Biography and Experience</h4>
     <div class="address-details" id="BioExp">
     <h5 id="shortbioh">Short Bio</h5>
     <p id="shortbio" class="lblRight"></p>
     <h5 id="bioAttachhead">Bio Attachment(s)</h5>
     <div class="bio-attachment-section" id="bioAttachment"></div>
     <div class="other-exp-view">
     <h5 class="other-exp-h">Other Experience Details</h5>
     <div class="exp">
     <div class="w-50" id="industryExpViewSec"><label id="IndustryExpl">Industries</label>
     <p id="IndustryExp" class="lblRight"></p>
     </div>
     <div class="w-50" id="languageExpViewSec"><label id="LanguageExpl">Languages</label>
     <p id="LanguageExp" class="lblRight"></p>
     </div>
     <div class="w-50" id="courseExpViewSec"><label id="SDGCoursel">Courses</label>
     <p id="SDGCourse" class="lblRight"></p>
     </div>
     <div class="w-50" id="softwareExpViewSec"><label id="SoftwareExpl">Software</label>
     <p id="SoftwareExp" class="lblRight"></p>
     </div>
     <div class="w-50" id="membershipViewSec"><label id="MembershipExpl">Memberships</label>
     <p id="MembershipExp" class="lblRight"></p>
     </div>
     <div class="w-50" id="specialKnowledgeViewSec"><label id="SpecialKnowledgel">Special Knowledge</label> 
     <p id="SpecialKnowledge" class="lblRight"></p>
     </div>
     </div>
     
     </div>
     <!--<div class="showAvail-view">
     <h4>Show Availability</h4>
     <label id="view-show-avail" class="lblRight"></label>
     </div>-->
     </div>
     </div>
     </div> 
     </div>
     <div id="DirectoryInformationEdit" class="edit-directory hide">
     <div class="d-flex">
     <div class="DInfo-left col-6">
     <div class="work-address">
     <h4>Office Affiliation</h4>
     <div class="address-details d-flex" id="editWorAddress">
     <label>Location</label>
     <div class="w-100"><select id="workLocationDD"></select></div>
     </div>
     <div class="Location-Addresses d-flex">
     <label>Location Address</label>
     <div class="address-details lblRight w-100" id="EditedAddressDetails">

     </div>
     </div>
     </div>
     <div class="staff-function-edit-info">
     <div class="d-flex">
     <label>Title/Staff Function :</label>
     <div class="w-100"><select id="StaffFunctionEdit"></select></div>
     </div>
     </div>
     <div class="staff-affiliates-edit-info">
     <div class="d-flex">
     <label>Staff Affiliates</label>
     <div class="w-100"><select id="StaffAffiliatesEdit"></select></div>
     </div>
     </div>
     <div class="assisstant-info">
     <h4>Assisstant</h4>
     <div class="assisstant-name d-flex">
     <label>Name</label>
     <div class="w-100"><div id="peoplepickerText" title="APickerField" style="display:none"></div><select id="drpStaffAssitant"></select></div>

     </div>
     </div>
     <div class="personal-info">
     <h4>Personal Info</h4>
     <div class="address-details" id="PersonaInfo">
     <div class="d-flex"><label>Home Address:</label><div class="w-100"><input type="text" id="PAddLineE"></div></div>
     <div class="d-flex"><label>City:</label><div class="w-100"><input type="text" id="PAddCityE"></div></div>
     <div class="d-flex"><label>State:</label><div class="w-100"><input type="text" id="PAddStateE"></div></div>
     <div class="d-flex"><label>Postal Code:</label><div class="w-100"><input type="text" id="PAddPCodeE"></div></div>
     <div class="d-flex"><label>Country:</label><div class="w-100"><input type="text" id="PAddCountryE"></div></div>
     <div class="d-flex w-100"><label>Nationality:</label><div class="w-100"><input type="text" id="citizenshipE"></div></div>
     <div class="d-flex"><label>Significant Other :</label><div class="w-100"><textarea id="significantOther"></textarea></div></div>
     <div class="d-flex"><label>Children :</label><div class="w-100"><textarea id="children"></textarea></div></div>
     </div>
     </div>

     <div class="contact-info">
     <h4>Contact Info</h4>
     <div class="address-details" id="ContactInfo">
     <div class="d-flex"><label>Personal Mail :</label><div class="w-100"><input type="text" id="personalmailID"></div></div>
     <div class="d-flex"><label>Mobile No :</label><div class="w-100" id ="mobileNoSec"><div class="d-flex mobNumbers"><select class="mobNoCode"></select><input type="text" class="mobNo" id="mobileno1"/><span class="addMobNo add-icon"></span></div></div></div>
     <div class="d-flex"><label>Home No :</label><div class="w-100" id="homeNoSec"><div class="d-flex homeNumbers"><select class="homeNoCode"></select><input type="text" class="homeno" id="homeno"/><span class="addHomeNo add-icon"></span></div></div></div>
     <div class="d-flex"><label>Emergency No :</label><div class="w-100" id="emergencyNoSec"><div class="d-flex emergencyNumbers"><select class="emergencyNoCode"></select><input type="text" class="emergencyno" id="emergencyno" /><span class="addEmergencyNo add-icon"></span></div></div></div>
     <div class="d-flex hide"><label>Office No :</label><div class="w-100" id="officeNoSec"><div class="d-flex officeNumbers"><select class="officeNoCode"></select><input type="text" class="officeno" id="officeno"/><span class="addOfficeNo add-icon"></span></div></div></div>

     <!--<div class="d-flex"><label>Significant Other :</label><div class="w-100"><textarea id="significantOther"></textarea></div></div>
     <div class="d-flex"><label>Children :</label><div class="w-100"><textarea id="children"></textarea></div></div>-->
     <div class="d-flex"><label>LinkedIn ID :</label><div class="w-100"><input type="text" id="linkedInID"></div></div>
     <div class="d-flex"><label>Skype :</label><div class="w-100"><input type="text" id="SkypeID"></div></div>
     </div>
     </div>



     <div class="StaffStatus">
     <h4>Staff Status</h4>
     <div class="d-flex w-100">
     <label>Status</label><div class="w-100"><select id="staffstatusDD"></select></div></div>
     <div id="workscheduleEdit">
     <div class="d-flex w-100 hide" id="workscheduleSec">
     <label>Work Schedule</label>
     <div class="w-100"><input type="text" id="workScheduleE"></div>
     </div>
     </div>
     </div>
     <div class="citizen-info">
     <div class="address-details" id="CitizenInfo">
     <!--<div class="d-flex w-100"><label>Nationality:</label><div class="w-100"><input type="text" id="citizenshipE"></div></div>-->
     </div>
     </div>
     </div>

     <div class="DInfo-right col-6">
     <h4 class="BRHead" style="display:none">Billing Rate</h4>
     <div class="user-billing-rates hide">
     
     <div id="BillingRateDetailsView" class="hide">
     <div id="BillingRateDetails">
     <div class="billing-rates"><label>USD Daily Rate</label><div class="usd-daily-rate" id="UsdDailyRate"></div></div>
     <div class="billing-rates"><label>USD Hourly Rate</label><div class="usd-hourly-rate" id="UsdHourlyRate"></div></div>
     <div class="billing-rates"><label>EUR Daily Rate</label><div class="eur-daily-rate" id="EURDailyRate"></div></div>
     <div class="billing-rates"><label>EUR Hourly Rate</label><div class="eur-hourly-rate" id="EURHourlyRate"></div></div>
     <div class="billing-effective-date"><label>Effective Date</label><div class="effective-date" id="EffectiveDate"></div></div>
     <div class="billing-rates"><label>Comments</label><div class="w-100"><textarea class="Billing-comments" id="BillingRateComments"></textarea></div></div>
     </div>
     </div>

     <div id="BillingRateDetailsEdit" class="hide">
     <div id="BillingRateDetails">
     <div class="billing-rates"><label>USD Daily Rate</label><div class="usd-daily-rate"></div><input type="number" id="USDDailyEdit"/></div>
     <div class="billing-rates"><label>USD Hourly Rate</label><div class="usd-hourly-rate"></div><input type="number" id="USDHourlyEdit" disabled/></div>
     <div class="billing-rates"><label>EUR Daily Rate</label><div class="eur-daily-rate"></div><input type="number" id="EURDailyEdit"/></div>
     <div class="billing-rates"><label>EUR Hourly Rate</label><div class="eur-hourly-rate"></div><input type="number" id="EURHourlyEdit" disabled/></div>
     <div class="billing-rates"><label>Other Currency</label><div class="eur-hourly-rate"></div><select id="othercurrDD"></select></div>
     <div class="billing-rates"><label>Daily Rate</label><div class="eur-hourly-rate"></div><input type="number" id="ODailyEdit"/></div>
     <div class="billing-rates"><label>Hourly Rate</label><div class="eur-hourly-rate"></div><input type="number" id="OHourlyEdit" disabled/></div>
     <div class="billing-effective-date"><label>Effective Date</label><div class="effective-date"><input type="date" id="EffectiveDateEdit"/></div></div>
     <div class="billing-rates"><label>Comments</label><div class="w-100"><textarea class="Billing-comments" id="BillingRateCommentsEdit"></textarea></div></div>
     </div></div>
     </div>
     <div class="Biography-Experience">
     <h4>Biography and Experience</h4>
     <div class="address-details" id="BioExp">
     <h5>Short Bio</h5>
     <div><textarea id="Eshortbio"></textarea></div>
     <h5>Bio Attachment(s)</h5>
     <div class="bio-attachment-section" id="bioAttachment">
     <div class="custom-file">
<input type="file" name="myFile" id="BioAttachEdit" multiple class="custom-file-input">
<label class="custom-file-label" for="BioAttachEdit">Choose File</label>
</div>
<div class="quantityFilesContainer quantityFilesContainer-static" id="filesfromfolder"></div>
<div class="quantityFilesContainer quantityFilesContainer-static" id="otherAttachmentFiles"></div>

     </div>
     <div class="other-exp">
     <h5>Other Experience Details</h5>
     <div class="exp">
     <div class=""><label>Industries</label>
     <div><textarea id="EIndustry"></textarea></div>
     </div>
     <div class=""><label>Languages</label>
     <div><textarea id="ELanguage"></textarea></div>
     </div>
     </div>
     <div class="exp">
     <div class=""><label>Courses</label>
     <div><textarea id="ESDGCourse"></textarea></div>
     </div>
     <div class=""><label>Software</label>
     <div><textarea id="ESoftwarExp"></textarea></div>
     </div>
     </div>
     <div class="exp">
     <div class=""><label>Memberships</label>
     <div><textarea id="EMembership"></textarea></div>
     </div>
     <div class=""><label>Special Knowledge</label>
     <div><textarea id="ESKnowledge"></textarea></div>
     </div>
     </div>
     </div>
     
     <div class="show-availablity-sec">
     <h5>Show Availability</h5>
      <label class="switch">
            <input type="checkbox" id="show-availability">   
            <span class="slider round"></span>
          </label>
      </div>
     </div>
     </div>
     </div>

     </div>
     <div class="btn-section">
     <button class="btn btn-cancel" id="BtnCancel">Cancel</button>
     <button class="btn btn-submit" id="BtnSubmit">Submit</button>
     </div>
     </div>
   </div>
   <div id="menu1" class="tab-pane fade">
     <div class="view-availability">
     <div class="availability-btn-section">
     <button class="btn btn-add-project"  data-toggle="modal" data-target="#addprojectmodal">Add Project</button>
     <div class="todayAvail"></div>
     <div class="NextAvail"></div>
     </div> 

     <!--<div class='FilterTableforProjecttype'>
    <div class="serchdivforProjecttype">
    <label>ProjectType:</label>
    <select id="drpforProjecttype">
    <option value="Select">Select</option>
    </select>
    </div>
    </div>-->

     <div class="availability-table-section">
     <table id="UserAvailabilityTable">
     <thead>
     <tr>
     <th>Project Type</th>
     <th class="w100">Project Name</th>
     <th>Start Date</th>
     <th>End Date</th>
     <th>% Time</th>
     <th class="w100">Comments</th>
     <th>Action</th>
     </tr>
     </thead>
     <tbody id="UserAvailabilityTbody">
     </tbody >
     </table>
     </div>
     </div> 


     <div class="modal fade" id="addprojectmodal" tabindex="-1" role="dialog" aria-labelledby="addprojectmodalLabel" aria-hidden="true">
 <div class="modal-dialog" role="document">
   <div class="modal-content">
     <div class="modal-header">
       <h5 class="modal-title" id="exampleModalLabel">Add Project</h5>

     </div>
     <div class="modal-body add-project-modal">
     <div class="d-flex" id="add-basic-project">
     <div class="d-flex col-6"><label>Project Type</label><div class="w-100"><select id="projecttypeDD"><option value="sample">Sample</option></select></div></div>
     <div class="d-flex col-6"><label id="pName-lbl">Project Name</label><div class="w-100"><input type="text" id="projectName" /></div></div>
     </div>
     <div class="d-flex" id="add-basic-date">
     <div class="d-flex col-6"><label>Start Date</label><div class="w-100"><input type="date" id="projectStartDate" /></div></div>
       <div class="d-flex col-6"><label>End Date</label><div class="w-100"><input type="date" id="projectEndDate" /></div></div>
     </div>
       <div class="d-flex" id="add-percent-project">
       <div class="d-flex col-6"><div id="percentageDiv" class="d-flex w-100"><label id="percent-label">% Time on Project</label><div class="w-100"><input type="number" id="projectPercent" /></div></div></div>

       </div>

       <div class="d-flex" id="add-client-project">
       <div class="d-flex col-6"><label>Client</label><div class="w-100"><input type="text" id="client" /></div></div>
       <div class="d-flex col-6"><label>Project Code</label><div class="w-100"><input type="text" id="projectCode" /></div></div>
       </div>

       <div class="d-flex" id="add-practice-project">
       <div class="d-flex col-6"><label>Practice Area</label><div class="w-100"><select id="practiceAreaDD"><option value="sample">Sample</option></select></div></div>
       <div class="d-flex col-6"><label>Project Location</label><div class="w-100"><input type="text" id="ProjectLocation" /></div></div>
       </div>
       <div class="d-flex" >
       <div class="d-flex col-6"><div id="OtherPracticeAreaDiv" class="d-flex w-100 hide"><label>Others</label><div class="w-100"><input type="text" id="OtherPracticeArea" /></div></div></div>

       </div>
       <div class="d-flex" id="add-cmt-project">
       <div class="d-flex col-6" id="projectAvailNotes-sec"><label>Availability Notes</label><div class="w-100"><textarea id="projectAvailNotes" ></textarea></div></div>
       <div class="d-flex col-6" id="Projectcomments-sec"><label>Comments</label><div class="w-100"><textarea id="Projectcomments"></textarea></div></div></div>
     </div>
     <div class="modal-footer">
       <button type="button" class="btn btn-cancel" data-dismiss="modal" id="closeModal">Close</button>
       <button type="button" class="btn btn-submit" id="add-availability">Submit</button>
     </div>
   </div>
 </div>
</div>


   </div>
   </div>
   </div>
    </div>
    </div>
    </div>
    </div>
    `;


      $(".nav-items").click(()=>{	
      $("#availabilityTab").removeClass("active")	
      $("#home").addClass("in active");	
      $("#menu1").removeClass("in active")	
      $("#directoryInfoTab").addClass("active")
    }); 

    $('.nav-tabs').click((e)=>{
      var slctEvent = e.target.id;
      if(slctEvent=="aDirectory")
      localStorage.setItem("SelectedTab",slctEvent)
      else if(slctEvent=="aAvailablity")
      localStorage.setItem("SelectedTab",slctEvent)
    });

    const username = document.querySelectorAll(".usernametag");
    const usernamelast = document.querySelectorAll(".usernametag-last");
    const userpage = document.querySelector(".user-profile-page");
    const tableSection = document.querySelector(".sdh-employee");
    const viewDir = document.querySelector(".view-directory");
    const editDir = document.querySelector(".edit-directory");
    const editbtn = document.querySelector(".btn-edit");

    // ! Side Nav Click Action
   {

    $(".clsToggleCollapse").click(function () {
      $(".clsCollapse").each(function () {
        $(this).removeClass("in").attr("style", "");
      });
      $(this).next("div").addClass("in");
    });

    onLoadData();
    getGroups();
    ActiveSwitch();


    $(".SDHEmployee").click(() => {
      SelectedUserProfile = [];
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };

        bindEmpTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindEmpTable(options);
      }
    });
    
    $(".OutsidConsultant").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");
      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindOutTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindOutTable(options);
      }
    });

    $(".SDHAffiliates").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");
      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindAffTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindAffTable(options);
      }
    });

    $(".SDHAlumini").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");
      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindAlumTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindAlumTable(options);
      }
    });


    $(".SDHShowAll").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");
      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindAllDetailTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindAllDetailTable(options);
      }
    });


    $(".SDGOfficeInfo").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");
      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindOfficeTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindOfficeTable(options);
      }
    });


  }
    // Employee Filters
    $(".sdhLocgrouping").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
      SdhEmpTableRowGrouping(4, "SdhEmpTable", bindEmpTable);
    });

    $(".sdhTitlgrouping").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
      SdhEmpTableRowGrouping(6, "SdhEmpTable", bindEmpTable);
    });



    


    $(".sdhAssistantgrouping").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
      SdhEmpTableRowGrouping(7, "SdhEmpTable", bindEmpTable);
    });

    $(".sdhfirstnamesort").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
      var options = {
        destroy: true,
        order: [[1, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindEmpTable(options);
    });

    $(".sdhlastnamesort").click(() => {
      $(".usernametag").hide();
      $(".usernametag-last").show();
     // $(".usernametag").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
      var options = {
        destroy: true,
        order: [[2, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindEmpTable(options);
    });

    //OutSideConsultant
    $(".OutConslastnamesort").click(() => {
      $(".usernametag").hide();
      $(".usernametag-last").show();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
      var options = {
        destroy: true,
        order: [[2, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindOutTable(options);
    });

    $(".OutConsFirstnamesort").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
      var options = {
        destroy: true,
        order: [[1, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindOutTable(options);
    });

    $(".OutConsLocgrouping").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      SdhEmpTableRowGrouping(4, "SdhOutsideTable", bindOutTable);
    });

    $(".OutConsStaffgrouping").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      SdhEmpTableRowGrouping(6, "SdhOutsideTable", bindOutTable);
    });

    // Affliates
    $(".Afflastnamesort").click(() => {
      $(".usernametag").hide();
      $(".usernametag-last").show();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[2, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindAffTable(options);
    });

    $(".AffFirstnamesort").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[1, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindAffTable(options);
    });

    // Allumni
    $(".SDHAlumniLastName").click(() => {
      $(".usernametag").hide();
      $(".usernametag-last").show();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[2, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindAlumTable(options);
    });

    $(".SDHAlumniFirstName").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[1, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindAlumTable(options);
    });

    $(".SDHAlumniOffice").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      SdhEmpTableRowGrouping(4, "SdhAllumniTable", bindAlumTable);
    });

    // All Users
    $(".SDHShowAllLastName").click(() => {
      $(".usernametag").hide();
      $(".usernametag-last").show();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[2, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindAllDetailTable(options);
    });

    $(".SDHShowAllFirstName").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[1, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindAllDetailTable(options);
    });

    $(".StaffAvailability").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[0, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
      };
      bindAllDetailTable(options);
    });

    $(".SDGBillingRate").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[0, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
        columnDefs : [
          //hide the second & fourth column
          { 'visible': false, 'targets': [5,6] }
        ]
      };
      bindBillingRateTable(options);
    });

    $(".SDGBillingRateTitle").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[1, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
        columnDefs : [
          //hide the second & fourth column
          { 'visible': false, 'targets': [5,6] }
        ]
      };
      bindBillingRateTable(options);
    });

    $(".SDGBillingRateLastName").click(() => {
      $(".usernametag").hide();
      $(".usernametag-last").show();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[6, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
        columnDefs : [
          //hide the second & fourth column
          { 'visible': false, 'targets': [5,6] }
        ]
      };
      bindBillingRateTable(options);
    });

    $(".SDGBillingRateFirstName").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

        }
        if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        }
      var options = {
        destroy: true,
        order: [[5, "asc"]],
        language: {
          "emptyTable": "No data available"
        },
        lengthMenu: [50, 100],
        columnDefs : [
          //hide the second & fourth column
          { 'visible': false, 'targets': [5,6] }
        ]
      };
      bindBillingRateTable(options);
    });
    
    $(".SDGOfficeInfoFirstName").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");
      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
        var options = {
          destroy: true,
          order: [[1, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindOfficeTable(options);
    });
    
    $(".SDGOfficeInfoLastName").click(() => {
      $(".usernametag").hide();
      $(".usernametag-last").show();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");
      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
        var options = {
          destroy: true,
          order: [[2, "asc"]],
          language: {
            "emptyTable": "No data available"
          },
          lengthMenu: [50, 100],
        };
        bindOfficeTable(options);
    });

    $(".staffAvailOffice").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
      SdhEmpTableRowGrouping(1,"StaffAvailabilityTable",bindStaffAvailTable);
    });

    $(".staffAvailStaff").click(() => {
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");

      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
      SdhEmpTableRowGrouping(2 ,"StaffAvailabilityTable",bindStaffAvailTable);
    });

    $(".staffAvailAffiliation").click(()=>{
      $(".usernametag").show();
      $(".usernametag-last").hide();
      if (viewDir.classList["contains"]("hide")) {
        viewDir.classList.remove("hide");
        editDir.classList.add("hide");
        editbtn.classList.remove("hide");
      }
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
      }
      SdhEmpTableRowGrouping(3 ,"StaffAvailabilityTable",bindStaffAvailTable);


    });

    $("#btnEdit").click(() => {
      editFunction();
    });

    $("#BtnSubmit").click(() => {
      editsubmitFunction();
    });

    $("#BtnCancel").click(() => {
      editcancelFunction();
    });

    $("#add-availability").click(()=>{
      if(mandatoryforaddaction())
    {
      if(AvailEditFlag){
        availUpdateFunc();
      }else{
        availSubmitFunc();
      }
    }
      else{
        console.log("All fileds not filled");
      }
    });

    $(document).on("change", "#BioAttachEdit", function () {
      if ($(this)[0].files.length > 0) {
        for (let index = 0; index < $(this)[0].files.length; index++) {
          const file = $("#BioAttachEdit")[0]["files"][index];
          // if (ValidateSingleInput($("#others")[0])) {
          bioAttachArr.push(file);
          $("#otherAttachmentFiles").append(
            '<div class="quantityFiles">' +
              "<span class=upload-filename>" +
              file.name +
              "</span>" +
              "<a filename='" +
              file.name +
              "' class=clsRemove href='#'>x</a></div>"
          );
          // }
        }
        $(this).val("");
        $(this).parent().find("label").text("Choose File");
      }
    });

    $(document).on("click", ".clsRemove", function () {
      //var filename=$(this).attr('filename');
      var filename = $(this).parent().children()[0].innerText;
      removeSelectedfile(filename);
      $(this).parent().remove();
    });


    $(document).on("click", ".remove-icon", function () {
      $(this).parent().remove();
    });

    $(document).on("click",".action-delete",(e)=>{
      let AItemID = e.currentTarget.getAttribute("data-id");
      alertify.confirm("Are you sure want to delete",
  function(){
    removeAvailProject(parseInt(AItemID));
      e.currentTarget.parentElement.parentElement.parentElement.remove()
  },
  function(){
    
  });
      
    });

    $(document).on("change", "#staffstatusDD", function(){
      if ($("#staffstatusDD").val() == "Part-time") {
        $("#workscheduleSec").removeClass("hide");
      } else if($("#staffstatusDD").val() == "Full-time"){
        $("#workscheduleSec").addClass("hide");
        $("#workscheduleSec").val("");
      }
    });

    $(document).on("change","#workLocationDD",function(){
      $("#EditedAddressDetails").html(OfficeAddArr.filter(
        (add) => $("#workLocationDD").val() == add.OfficeID
      )[0].OfficeFullAdd);
    });

    $(document).on("change","#projecttypeDD",()=>{
      if( $("#projecttypeDD").val() == "Marketing"|| $("#projecttypeDD").val() == "Overhead"){
        $('#pName-lbl').text("Title");
        $('#percent-label').text("Percentage");
        $('#projectName').val($("#projecttypeDD").val());
        $('#projectName').prop("disabled",false);
        $("#practiceAreaDD").val("Select");
        $("#OtherPracticeAreaDiv").addClass('hide')
      $('#add-basic-project,#add-basic-date,#add-percent-project,#Projectcomments-sec').show();
      $('#add-client-project,#add-practice-project,#projectAvailNotes-sec').hide();
      

      }
      else if($("#projecttypeDD").val() == "Vacation")
      {
        $('#pName-lbl').text("Title");
        $('#percent-label').text("Percentage");
        $('#projectName').val($("#projecttypeDD").val());
        $('#projectName').prop("disabled",true);
        $("#practiceAreaDD").val("Select");
        $("#OtherPracticeAreaDiv").addClass('hide')
      $('#add-basic-project,#add-basic-date,#add-percent-project,#Projectcomments-sec').show();
      $('#add-client-project,#add-practice-project,#projectAvailNotes-sec').hide();
        
      }else{
        $('#pName-lbl').text("Project Name");
        $('#percent-label').text("% Time on Project");
        $("#practiceAreaDD").val("Select");
        $('#projectName').val("");
        $("#OtherPracticeArea").val("");
        $('#projectName').prop("disabled",false);
        $("#OtherPracticeAreaDiv").addClass('hide')
        $('#add-basic-project,#add-basic-date,#add-percent-project,#Projectcomments-sec').show();
        $('#add-client-project,#add-practice-project,#projectAvailNotes-sec').show();
      }
    });

    $(document).on("change","#practiceAreaDD",()=>{
      if($("#practiceAreaDD").val() == "Others, please specify"){

          $("#OtherPracticeAreaDiv").removeClass('hide')

      }else{
        $("#OtherPracticeArea").val("");
          $("#OtherPracticeAreaDiv").addClass('hide')

      }
    });

    $(document).on("click","#editProjectAvailability",(e)=>{
      AvailEditFlag = true;
      let AEditItemID = e.currentTarget.getAttribute("data-id");
      AvailEditID = AEditItemID;
      $('#exampleModalLabel').empty();
      $('#exampleModalLabel').append("Edit Project");
      fillEditSection(AvailEditID);
    });

    $(document).on("click","#closeModal",()=>{
      $('#exampleModalLabel').empty();
      $('#exampleModalLabel').append("Add Project");
      AvailEditFlag = false;
      AvailEditID = 0;
      $("#projectName").val("")
      $("#projectStartDate").val("")
      $("#projectEndDate").val("")
      $("#projectPercent").val("")
      $("#practiceAreaDD").val("Select")
      $("#projecttypeDD").val("Select")
      $("#client").val("")
      $("#projectCode").val("")
      $("#ProjectLocation").val("")
      $("#OtherPracticeAreaDiv").addClass('hide')
      $("#OtherPracticeArea").val("")
      $("#projectAvailNotes").val("")
      $("#Projectcomments").val("")

      //$('#projectName').prop("disabled",false);
      $('#add-basic-project,#add-basic-date,#add-percent-project,#Projectcomments-sec').show();
      $('#add-client-project,#add-practice-project,#projectAvailNotes-sec').show();

      $('#pName-lbl').text("Project Name");
      $('#percent-label').text("% Time on Project");

    });

    $(document).on("click",".usernametag",(e)=>{
      var userName=e.target.id;
      var mainsideshow=$('.card .show').attr('id');
      var secSideShow = $('.card .show').next().attr('id');
      IsAdminStaff||(SelectedUserProfile[0].Usermail.toLowerCase()==currentMail.toLowerCase()&&IsgeneralStaff)?$('.btn-add-project').show():$('.btn-add-project').hide()


      localStorage.setItem("EName", userName);
      localStorage.setItem("MainSideNav", mainsideshow);
      localStorage.setItem("secSideShow", secSideShow);
      localStorage.setItem("SelectedTab", "aDirectory");

      if((SelectedUserProfile[0].Affiliation != "Employee" && SelectedUserProfile[0].Affiliation != "Outside Consultant")||!SelectedUserProfile[0].showAvailability){
        $("#menu1").addClass("hide");
        $("#availabilityTab").addClass("hide");
      }else{
        $("#menu1").removeClass("hide");
        $("#availabilityTab").removeClass("hide");
      }
    });

    $(document).on("click",".usernametag-last",(e)=>{
      var userName=e.target.id;
      var mainsideshow=$('.card .show').attr('id');
      var secSideShow = $('.card .show').next().attr('id');
      IsAdminStaff||(SelectedUserProfile[0].Usermail.toLowerCase()==currentMail.toLowerCase()&&IsgeneralStaff)?$('.btn-add-project').show():$('.btn-add-project').hide()


      localStorage.setItem("EName", userName);
      localStorage.setItem("MainSideNav", mainsideshow);
      localStorage.setItem("secSideShow", secSideShow);
      localStorage.setItem("SelectedTab", "aDirectory");

      if((SelectedUserProfile[0].Affiliation != "Employee" && SelectedUserProfile[0].Affiliation != "Outside Consultant")||!SelectedUserProfile[0].showAvailability){
        $("#menu1").addClass("hide");
        $("#availabilityTab").addClass("hide");
      }else{
        $("#menu1").removeClass("hide");
        $("#availabilityTab").removeClass("hide");
      }
    });
    
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
const onLoadData = async () => {
  $(".loader-section").show();
  let LocOptionHtml = "";
  let LocDDHtml ="<option value='Select'>Select</option>"
  let LocValueHtml ="<option value='Select'>Select</option>"
  let staffOptionHtml = "";
  let otherCurrHtml = "";
  let StaffFunHtml = "<option value='Select'>Select</option>";
  let StaffDDHtml = "<option value='Select'>Select</option>"
  let StaffAffHtml = "<option value='Select'>Select</option>";
  let AvailProjTypeHtml = "<option value='Select'>Select</option>";
  let AvailPracAreaDD = "<option value='Select'>Select</option>";

  let listLocation  = await sp.web.getList(listUrl + "SDGOfficeInfo").items.get(); 
  //let listLocation = await sp.web.getList(listUrl + "StaffDirectory").fields.filter("EntityPropertyName eq 'SDGOffice'").get();

  let listStaffStatus = await sp.web.getList(listUrl + "StaffDirectory").fields.filter("EntityPropertyName eq 'StaffStatus'").get();

  let listOtherCurr = await sp.web.getList(listUrl + "StaffDirectory").fields.filter("EntityPropertyName eq 'OtherCurrency'").get();

  // let CountryCode = await sp.web.getList(listUrl + "StaffDirectory").fields.filter("EntityPropertyName eq 'CountryCode'").get();

  let listStaffFunction = await sp.web.getList(listUrl + "StaffDirectory").fields.filter("EntityPropertyName eq 'stafffunction'").get();

  let listStaffAff = await sp.web.getList(listUrl + "StaffDirectory").fields.filter("EntityPropertyName eq 'SDGAffiliation'").get();

  let AvailProjectType = await sp.web.getList(listUrl + "SDGAvailability").fields.filter("EntityPropertyName eq 'ProjectType'").get();

  let AvailPracticeArea = await sp.web.getList(listUrl + "SDGAvailability").fields.filter("EntityPropertyName eq 'ProjectArea'").get();

//ProjectType-Load
  AvailProjectType[0]["Choices"].forEach((type) => {
    AvailProjTypeHtml += `<option value="${type}">${type}</option>`;
  });
  
//ProjectArea-Load
    AvailPracticeArea[0]["Choices"].forEach((Area) => {
      AvailPracAreaDD += `<option value="${Area}">${Area}</option>`;
  });

//Location-Load
  // listLocation[0]["Choices"].forEach((li) => {
  //   LocOptionHtml += `<option value="${li}">${li}</option>`;
  //   LocDDHtml +=`<option value="${li}">${li}</option>`;
  // });
  listLocation.forEach((li) => {
      // LocOptionHtml += `<option value="${li.ID}">${li.Office}</option>`;
      LocDDHtml +=`<option value="${li.ID}">${li.Office}</option>`;
      LocValueHtml += `<option value="${li.Office}">${li.Office}</option>`;
    });

//Status-Load
  listStaffStatus[0]["Choices"].forEach((stff) => {
    staffOptionHtml += `<option value="${stff}">${stff}</option>`;
  });
//OtherCurrency-Load
  listOtherCurr[0]["Choices"].forEach((curr) => {
    otherCurrHtml += `<option value="${curr}">${curr}</option>`;
  });
//CountryCode - Load
  // CountryCode[0]["Choices"].forEach((CCode) => {
  //   CCodeArr.push(CCode);
  //   CCodeHtml += `<option value="${CCode}">${CCode}</option>`;
  // });

//StaffFunction - Load
  listStaffFunction[0]["Choices"].forEach((func) => {
    StaffFunHtml += `<option value="${func}">${func}</option>`;
    StaffDDHtml += `<option value="${func}">${func}</option>`;
  });

//Affliation - Load
  listStaffAff[0]["Choices"].forEach((Aff) => {
    StaffAffHtml += `<option value="${Aff}">${Aff}</option>`;
  });

  $('#projecttypeDD').html(AvailProjTypeHtml);
 // $('#drpforProjecttype').html(AvailProjTypeHtml);
  $("#practiceAreaDD").html(AvailPracAreaDD);
  // $('#workLocationDD').html(LocOptionHtml);
  $('#staffstatusDD').html(staffOptionHtml);
  $('#othercurrDD').html(otherCurrHtml);
  $("#StaffFunctionEdit,#drpAffiliatesforBilling").html(StaffFunHtml);
  $("#StaffAffiliatesEdit,#drpStaffforBilling").html(StaffAffHtml);
  $(".mobNoCode,.homeNoCode,.emergencyNoCode,.officeNoCode").html(CCodeHtml);
  $("#drpTitleforEmployee,#drpTitleforOutside,#drpTitleforAffiliates,#drpTitleforAlumni,#drpTitleforAllPeople,#drpTitleforBilling").html(StaffDDHtml);
  $("#drpLocationforEmployee,#drpLocationforOutside,#drpLocationforAffiliates,#drpLocationforAlumni,#drpLocationforAllPeople").html(LocValueHtml); 

  ProfilePics = await sp.web.getFolderByServerRelativeUrl(`/sites/StaffDirectory/ProfilePictures`).files.select("*,listItemAllFields").expand("listItemAllFields").get();

  AllAvailabilityDetails = await sp.web.getList(listUrl + "SDGAvailability").items.select("*,UserName/Title,UserName/EMail,UserName/Id").orderBy('Modified', false).expand("UserName").top(5000).get();

  await sp.web.getList(listUrl + "StaffDirectory").items.select("*","User/EMail","User/Title","User/FirstName","User/LastName","User/JobTitle","User/UserName","Assistant/EMail","Assistant/Title","User/Id","SDGOfficeDetails/Office","SDGOfficeDetails/ID").expand("User,Assistant,SDGOfficeDetails").get().then((listitem: any) => {
    let tempArr = listitem.filter((l)=>l.SDGOfficeDetails != null)
        console.log(tempArr);
      listitem.forEach((li) => {
        // 
        let EndDateArr = [];
        let NextAvailDate = ""
        let modifiedArray =[];
        let fDateArr =[];
        let EndPercentage = 0
        // 
        
      var userPercentage=0;
      
      ///var userpic =   ProfilePics.filter((p)=>p.ListItemAllFields.UserName&& li.User.UserName ? p.ListItemAllFields.UserName.toLowerCase()==li.User.UserName.toLowerCase():"");

      var userpic =   ProfilePics.filter((p)=>p.ListItemAllFields.UserName&& li.UserEMail ? p.ListItemAllFields.UserName.toLowerCase()==li.UserEMail.toLowerCase():"");

      var datetoday:any=new Date(new Date().toLocaleDateString()+" 00:00")



      AllAvailabilityDetails.forEach((all)=>
      {
        
        ///all.UserName.EMail == li.User.EMail && new Date(new Date(all.StartDate).toLocaleDateString()+" 00:00")<=datetoday && new Date(new Date(all.EndDate).toLocaleDateString()+" 00:00")>=datetoday&&all.ProjectType!="Marketing"&&all.ProjectType!="Overhead"?userPercentage += parseInt(all.Percentage):userPercentage += 0;

        all.UserEMail == li.UserEMail && new Date(new Date(all.StartDate).toLocaleDateString()+" 00:00")<=datetoday && new Date(new Date(all.EndDate).toLocaleDateString()+" 00:00")>=datetoday&&all.ProjectType!="Marketing"&&all.ProjectType!="Overhead"?userPercentage += parseInt(all.Percentage):userPercentage += 0;

        
    var edate=new Date(new Date(all.EndDate).toLocaleDateString()+" 00:00");
    var todayDate =new Date(new Date().toLocaleDateString()+" 00:00");
    ///if(all.UserName.EMail == li.User.EMail && edate>=todayDate&&all.ProjectType!="Marketing"&&all.ProjectType!="Overhead")
    if(all.UserEMail == li.UserEMail && edate>=todayDate&&all.ProjectType!="Marketing"&&all.ProjectType!="Overhead")
    {
      EndDateArr.push(new Date(all.EndDate));
      modifiedArray.push(all);
    }
        
      });
      console.log(EndDateArr.sort(sortFunction));
       EndDateArr =EndDateArr.sort(sortFunction);
      if(EndDateArr.length>0)
      {
        for(let i=0;i<EndDateArr.length; i++)
        {
          fDateArr =[];
          var endPercentage=0
          var leastEndDate =EndDateArr[i];
           modifiedArray.forEach((cal)=>{
            var edate=new Date(new Date(cal.EndDate).toLocaleDateString()+" 00:00");
            if(new Date(edate).toLocaleDateString()==new Date(leastEndDate).toLocaleDateString())
            fDateArr.push(cal)
          });
          fDateArr.map((b)=>b.Percentage?endPercentage+=parseInt(b.Percentage):endPercentage+=0);
          if(endPercentage<100){

            // var nextDay = new Date(leastEndDate);
            // nextDay.setDate(nextDay.getDate() + 1);
            // NextAvailDate = nextDay.toLocaleDateString();
            EndPercentage = endPercentage;
            NextAvailDate = new Date(leastEndDate).toLocaleDateString();

            // $('.todayAvail').html(`<label>Next available on :${new Date(leastEndDate).toLocaleDateString()}(${100-endPercentage}% Available)</label>`);
            break;
          }
          else if(i==EndDateArr.length-1)
          {
            var tomorrow = new Date(leastEndDate);
            tomorrow.setDate(tomorrow.getDate() + 1);
            NextAvailDate = tomorrow.toLocaleDateString();
            // $('.todayAvail').html(`<label>Next available on :${tomorrow.toLocaleDateString()}(${100}% Available)</label>`);
            break;
          }
        }
  
      }
      else{
        
        NextAvailDate = new Date().toLocaleDateString();
        // NextAvailDate = new Date(leastEndDate).toLocaleDateString();
        // $('.todayAvail').html(`<label>Next available on :${new Date(leastEndDate).toLocaleDateString()}</label>`)
      }
      // console.log(NextAvailDate);
      UserDetails.push({
          
          /*Name: li.User.Title  ? li.User.Title : "",
          FirstName: li.User.FirstName  ? li.User.FirstName : "",
          LastName: li.User.LastName  ? li.User.LastName  : "",
          Usermail: li.User.UserName  ? li.User.UserName : li.User.EMail?li.User.EMail:"",
          UserId:li.User.Id,
          JobTitle: li.User.JobTitle  ? li.User.JobTitle : "Not Available",*/
          Name: li.UserTitle  ? li.UserTitle : "",
          FirstName: li.UserFirstName  ? li.UserFirstName : "",
          LastName: li.UserLastName  ? li.UserLastName  : "",
          Usermail: li.UserEMail?li.UserEMail:"",
          //UserId:li.User.Id,
          JobTitle: li.UserJobTitle  ? li.UserJobTitle : "Not Available",
          AssisstantName:li.AssisstantName  ? li.AssisstantName : "",

          UserPersonalMail: li.PersonalEmail  ? li.PersonalEmail : "",
          ///Assistant: li.Assistant  ? li.Assistant.Title : "",
          Assistant: li.AssisstantName  ? li.AssisstantName : "",
          AssistantMail: li.Assistant  ? li.Assistant.EMail : "",
          PhoneNumber: li.MobileNo  ? li.MobileNo : "",
          Location: li.SDGOfficeDetails !=null ?li.SDGOfficeDetails.Office  ? li.SDGOfficeDetails.Office : "":"",
          Locationid: li.SDGOfficeDetails !=null ?li.SDGOfficeDetails.ID  ? li.SDGOfficeDetails.ID : "":"",
          Title: li.stafffunction  ? li.stafffunction : "",
          Affiliation: li.SDGAffiliation  ? li.SDGAffiliation : "",
          HAddLine: li.HomeAddLine  ? li.HomeAddLine : "",
          HAddCity: li.HomeAddCity  ? li.HomeAddCity : "",
          HAddState: li.HomeAddState  ? li.HomeAddState : "",
          HAddPCode: li.HomeAddPCode  ? li.HomeAddPCode : "",
          HAddPCountry: li.HomeAddCountry  ? li.HomeAddCountry : "",
          ShortBio: li.ShortBio  ? li.ShortBio : "",
          Citizen: li.Citizenship  ? li.Citizenship : "",
          Industry: li.IndustryExp  ? li.IndustryExp : "",
          Language: li.LanguageExp  ?li.LanguageExp : "",
          SDGCourse: li.SDGCourses  ? li.SDGCourses : "",
          Software: li.SoftwareExp  ? li.SoftwareExp : "",
          Membership: li.Membership  ? li.Membership : "",
          SpecialKnowledge: li.SpecialKnowledge  ? li.SpecialKnowledge : "",
          USDDaily: li.USDDailyRate,
          USDHourly: li.USDHourlyRate,
          EURDaily: li.EURDailyRate,
          EURHourly: li.EURHourlyRate,
          OtherCurr: li.OtherCurrency, 
          OtherCurrDaily: li.ODailyRate,
          OtherCurrHourly: li.OHourlyRate,
          EffectiveDate: li.EffectiveDate ? li.EffectiveDate : "",
          BillingRateComments:li.BillingRateComments,
          StaffStatus: li.StaffStatus  ? li.StaffStatus : "",
          WorkSchedule: li.WorkingSchedule  ? li.WorkingSchedule : "",
          ItemID: li.ID  ? li.ID : "",
          LinkedInID: li.LinkedInLink  ? li.LinkedInLink : "",
          Skype:li.Skype?li.Skype : "",
          SignOther: li.signother  ? li.signother : "",
          Child: li.children  ? li.children : "",
          HomeNo: li.HomeNo  ? li.HomeNo : "",
          EmergencyNo: li.EmergencyNo  ? li.EmergencyNo : "",
          OfficeNo: li.OfficeNo  ? li.OfficeNo : "",
          ProfilePic:userpic.length>0?userpic[0].ServerRelativeUrl:"/sites/StaffDirectory/SiteImages/profile.png",
          Availability:userPercentage==0?100:100-userPercentage,
          showAvailability:li.ShowAvailability?true:false,
          NextAvailDate:NextAvailDate,
          EndPercentage:EndPercentage
        });
      });

     

      getTableData();
    });
    $(".loader-section").hide();   
  };
const ActiveSwitch = () => {
  let navItems = document.querySelectorAll(".nav-items");
  $('#collapseOne').addClass('in')
  navItems.forEach((li) => {
    li.addEventListener("click", (e) => {
      localStorage.setItem("EName", "");
      let activeClass = document.querySelectorAll(".nav-items");

      activeClass.forEach((activeC) => {
        activeC["classList"].remove("show");
      });

      let selectedOption = e.currentTarget;
      e.currentTarget["classList"].toggle("show");
      let activeTable = document.querySelectorAll(".oDataTable");
      activeTable.forEach((tables) => {
        if (!tables.classList.contains("hide")) {
          tables.classList.add("hide");
        }
        selectedOption["classList"].contains("SDHEmployee")
          ? $(".sdh-emp-table").removeClass("hide")
          : selectedOption["classList"].contains("OutsidConsultant")
          ? $(".sdh-outside-table").removeClass("hide")
          : selectedOption["classList"].contains("SDHAffiliates")
          ? $(".sdh-Affilate-table").removeClass("hide")
          : selectedOption["classList"].contains("SDHAlumini")
          ? $(".sdh-Allumni-table").removeClass("hide")
          : selectedOption["classList"].contains("SDHShowAll")
          ? $(".sdh-AllPeople-table").removeClass("hide")
          : selectedOption["classList"].contains("SDGOfficeInfo")
          ? $(".sdgofficeinfotable").removeClass("hide")
          : selectedOption["classList"].contains("SDGBillingRate")
          ? $(".sdgbillingrateTable").removeClass("hide")
          :  selectedOption["classList"].contains("StaffAvailability")?$(".StaffAvailabilityTable").removeClass("hide"):"";
      });
    });
  });
};
async function getTableData() {
  let OfficeTable = "";
  let EmpTable = "";
  let OutTable = "";
  let AffTable = "";
  let AlumTable = "";
  let AllDetailsTable = "";
  let BillingRateTable = "";

  let AvailHtml = "";
  let AssDDHtml ="<option value='Select'>Select</option>"
  let drpArray=[];
  let drpAss=[];
  let arrAssitant=[];
   OfficeDetails = await sp.web.getList(listUrl + "SDGOfficeInfo").items.get(); 

  OfficeDetails.forEach((oDetail) => {
    OfficeTable += `<tr><td>${oDetail.Office}</td><td>${
      oDetail.Phone != "null" ? oDetail.Phone.split("^").join("</br>") : ""
    }</td><td>${
      oDetail.Address != "null" ? oDetail.Address.split("^").join("</br>") : ""
    }</td></tr>`;
  });



//Availablity Table -Load 

    UserDetails.forEach((avli)=>{
      // console.log(avli);
      
      if(avli.Affiliation!="Alumni"&&avli.Affiliation!="Affiliate"&&avli.showAvailability!=false)
      {
        AvailHtml+=`<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${avli.Usermail}><img src="${avli.ProfilePic}" width="30" height="30" />${avli.Name}</div><div class="HUserDetails">
        <img src="${avli.ProfilePic}" class="userimg"/>
        <div class="user-name">${avli.Name}</div>
        <div class="user-JTitle">${avli.Usermail}</div>
        <div class="user-avail-title">Availability</div>
        <div class="user-percent">${avli.Availability}%</div>
        </div></div></td>
        <td>${avli.Location}</td>
        <td>${avli.Title}</td>
        <td>${avli.Affiliation}</td>
        <td> 
        <div class="d-flex align-item-center  availabilityProgress">
        
        <div class="nextAvailDate">${avli.Availability == 0 ?`${avli.NextAvailDate} (${100 - avli.EndPercentage} %)` :""}</div>
        
        <div class="availability-progress-bar" style="">
        <div class="progress-value" style="height:100%;width:${avli.Availability}%; background: ${avli.Availability <= 50? "#f01616":"#45b345"}"></div>
        </div>
        <span style="color:${avli.Availability <= 50? "#000000":"#000000"}">${avli.Availability}%</span></div>
        </td></tr>`
      }

    })
    $('#StaffAvailabilityTbody').html(AvailHtml);
 
  UserDetails.forEach((details) => {
    let ViewPhoneNumber = details.PhoneNumber.split("^");
    ViewPhoneNumber.pop(); 
    if (details.Affiliation == "Employee") {

      if(details.LastName!=""){

      EmpTable += `<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.Name}</div>

      <div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.LastName}, ${details.FirstName}</div>
      
      <div class="HUserDetails">
      <img src="${details.ProfilePic}" class="userimg"/>
      <div class="user-name">${details.Name}</div>
      <div class="user-JTitle">${details.Usermail}</div>
      <div class="user-avail-title">Availability</div>
      <div class="user-percent">${details.Availability}%</div>
      </div></div></td><td>${
        details.FirstName
      }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
        "<br>"
        )}</td><td>${
          details.Location == "" ||details.Location == null
          ? "Not Available"
          : `${details.Location}`
          }</td><td>${
          details.JobTitle == "" ||details.JobTitle == null
          ? "Not Available"
          : `${details.JobTitle}`
          }</td><td>${
          details.Title
        }</td><td>${
          details.Assistant == "" ||details.Assistant == null
          ? "Not Available"
          : `${details.Assistant}`
        }</td></tr>`;
      }

      else{
        EmpTable += `<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.Name}</div>

      <div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.FirstName}</div>
      
      <div class="HUserDetails">
      <img src="${details.ProfilePic}" class="userimg"/>
      <div class="user-name">${details.Name}</div>
      <div class="user-JTitle">${details.Usermail}</div>
      <div class="user-avail-title">Availability</div>
      <div class="user-percent">${details.Availability}%</div>
      </div></div></td><td>${
        details.FirstName
      }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
        "<br>"
        )}</td><td>${
          details.Location == "" ||details.Location == null
          ? "Not Available"
          : `${details.Location}`
          }</td><td>${
          details.JobTitle == "" ||details.JobTitle == null
          ? "Not Available"
          : `${details.JobTitle}`
          }</td><td>${
          details.Title
        }</td><td>${
          details.Assistant == "" ||details.Assistant == null
          ? "Not Available"
          : `${details.Assistant}`
        }</td></tr>`;
      }

    }
    
    if (details.Affiliation == "Outside Consultant") {

       if(details.LastName!=""){
     
      OutTable += `<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.Name}</div>
      
      <div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.LastName}, ${details.FirstName}</div>

      <div class="HUserDetails">
      <img src="${details.ProfilePic}" class="userimg"/>
      <div class="user-name">${details.Name}</div>
      <div class="user-JTitle">${details.Usermail}</div>
      <div class="user-avail-title">Availability</div>
      <div class="user-percent">${details.Availability}%</div>
      </div></div></td><td>${
        details.FirstName
      }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
        "<br>"
        )}</td><td>${
          details.Location == "" ||details.Location == null
          ? "Not Available"
          : `${details.Location}`
          }</td><td>${
          details.JobTitle == "" ||details.JobTitle == null
          ? "Not Available"
          : `${details.JobTitle}`
          }</td><td>${
          details.Title == "" ||details.Title == null
          ? "Not Available"
          : `${details.Title}`
          }</td><td>${
          details.Assistant == "" ||details.Assistant == null
          ? "Not Available"
          : `${details.Assistant}`
        }</td></tr>`;
      }

     else{
     
        OutTable += `<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.Name}</div>
        
        <div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.FirstName}</div>
  
        <div class="HUserDetails">
        <img src="${details.ProfilePic}" class="userimg"/>
        <div class="user-name">${details.Name}</div>
        <div class="user-JTitle">${details.Usermail}</div>
        <div class="user-avail-title">Availability</div>
        <div class="user-percent">${details.Availability}%</div>
        </div></div></td><td>${
          details.FirstName
        }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
          "<br>"
          )}</td><td>${
            details.Location == "" ||details.Location == null
            ? "Not Available"
            : `${details.Location}`
            }</td><td>${
            details.JobTitle == "" ||details.JobTitle == null
            ? "Not Available"
            : `${details.JobTitle}`
            }</td><td>${
            details.Title == "" ||details.Title == null
            ? "Not Available"
            : `${details.Title}`
            }</td><td>${
            details.Assistant == "" ||details.Assistant == null
            ? "Not Available"
            : `${details.Assistant}`
          }</td></tr>`;
        }

    }

    if (details.Affiliation == "Affiliate") {

      if(details.LastName!=""){

      AffTable += `<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.Name}</div>
      
      <div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.LastName}, ${details.FirstName}</div>

      </div></td><td>${
        details.FirstName
      }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
        "<br>"
        )}</td><td>${
          details.Location == "" ||details.Location == null
          ? "Not Available"
          : `${details.Location}`
          }</td><td>${
          details.JobTitle == "" ||details.JobTitle == null
          ? "Not Available"
          : `${details.JobTitle}`
          }</td><td>${
          details.Title == "" ||details.Title == null
          ? "Not Available"
          : `${details.Title}`
          }</td><td>${
          details.Assistant == "" ||details.Assistant == null
          ? "Not Available"
          : `${details.Assistant}`
        }</td></tr>`;
      }

      else{
        AffTable += `<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.Name}</div>
      
      <div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.FirstName}</div>

      </div></td><td>${
        details.FirstName
      }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
        "<br>"
        )}</td><td>${
          details.Location == "" ||details.Location == null
          ? "Not Available"
          : `${details.Location}`
          }</td><td>${
          details.JobTitle == "" ||details.JobTitle == null
          ? "Not Available"
          : `${details.JobTitle}`
          }</td><td>${
          details.Title == "" ||details.Title == null
          ? "Not Available"
          : `${details.Title}`
          }</td><td>${
          details.Assistant == "" ||details.Assistant == null
          ? "Not Available"
          : `${details.Assistant}`
        }</td></tr>`;
      }

    }
    if (details.Affiliation == "Alumni") {

      if(details.LastName!=""){

      AlumTable += `<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.Name}</div>
      
      <div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.LastName}, ${details.FirstName}</div>

      </div></td><td>${
        details.FirstName
      }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
        "<br>"
        )}</td><td>${
          details.Location == "" ||details.Location == null
          ? "Not Available"
          : `${details.Location}`
          }</td><td>${
          details.JobTitle == "" ||details.JobTitle == null
          ? "Not Available"
          : `${details.JobTitle}`
          }</td><td>${
          details.Title == "" ||details.Title == null
          ? "Not Available"
          : `${details.Title}`
          }</td><td>${
          details.Assistant == "" ||details.Assistant == null
          ? "Not Available"
          : `${details.Assistant}`
        }</td></tr>`;
      }

      else{
        AlumTable += `<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.Name}</div>
      
        <div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.FirstName}</div>
  
        </div></td><td>${
          details.FirstName
        }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
          "<br>"
          )}</td><td>${
            details.Location == "" ||details.Location == null
            ? "Not Available"
            : `${details.Location}`
            }</td><td>${
            details.JobTitle == "" ||details.JobTitle == null
            ? "Not Available"
            : `${details.JobTitle}`
            }</td><td>${
            details.Title == "" ||details.Title == null
            ? "Not Available"
            : `${details.Title}`
            }</td><td>${
            details.Assistant == "" ||details.Assistant == null
            ? "Not Available"
            : `${details.Assistant}`
          }</td></tr>`;
      }

    }
    AllDetailsTable += `<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.Name}</div>`

    if(details.LastName!=""){
    
      AllDetailsTable += `<div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.LastName}, ${details.FirstName}</div>`

    }

    else{
      AllDetailsTable += `<div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.FirstName}</div>`

    }

    
    if(details.Affiliation!="Alumni"&&details.Affiliation!="Affiliate")
    {
      AllDetailsTable +=  `<div class="HUserDetails">
      <img src="${details.ProfilePic}" class="userimg"/>
      <div class="user-name">${details.Name}</div>
      <div class="user-JTitle">${details.Usermail}</div>
      <div class="user-avail-title">Availability</div>
      <div class="user-percent">${details.Availability}%</div>
      </div>`
    }
 
    AllDetailsTable +=   `</div></td><td>${
      details.FirstName
    }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
      "<br>"
      )}</td><td>${
        details.Location == "" ||details.Location == null
        ? "Not Available"
        : `${details.Location}`
        }</td><td>${
        details.JobTitle == "" ||details.JobTitle == null
        ? "Not Available"
        : `${details.JobTitle}`
        }</td><td>${
        details.Title == "" ||details.Title == null
        ? "Not Available"
        : `${details.Title}`
        }</td><td>${
          details.Assistant == "" ||details.Assistant == null
          ? "Not Available"
          : `${details.Assistant}`
        }</td></tr>`;
        if(details.Assistant)
        {
          drpArray.push(details.Assistant);
        }
        if(details.Name)
        {
          arrAssitant.push(details.Name);
        }
        
        BillingRateTable += `<tr><td class="user-details-td"><div class="user-hover-details"><div class="usernametag" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.Name}</div>`

        if(details.LastName!=""){

          BillingRateTable += `<div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.LastName}, ${details.FirstName}</div>`

        }
        else{
          BillingRateTable +=  `<div class="usernametag-last" style="display: none" id=${details.Usermail}><img src="${details.ProfilePic}" width="30" height="30" />${details.FirstName}</div>`
        }
        
        
    if(details.Affiliation!="Alumni"&&details.Affiliation!="Affiliate")
    {
      BillingRateTable +=  `<div class="HUserDetails">
      <img src="${details.ProfilePic}" class="userimg"/>
      <div class="user-name">${details.Name}</div>
      <div class="user-JTitle">${details.Usermail}</div>
      <div class="user-avail-title">Availability</div>
      <div class="user-percent">${details.Availability}%</div>
      </div>`
    }
 
    BillingRateTable +=   `</div></td><td>${
      details.Title
    }</td><td><div>${
      details.USDDaily == "" || details.USDDaily == null
        ? ""
        : `USD: ${details.USDDaily}`
    }</div><div>${
      details.EURDaily == "" || details.EURDaily == null
        ? ""
        : `EUR: ${details.EURDaily}`
    }</div><div>${
      details.OtherCurrDaily == "" || details.OtherCurrDaily == null
        ? ""
        : `${details.OtherCurr}: ${details.OtherCurrDaily}`
    }</div></td><td><div>${
      details.USDDaily == "" || details.USDDaily == null
        ? ""
        : `USD: ${details.USDDaily/8}`
    }</div><div>${
      details.EURDaily == "" || details.EURDaily == null
        ? ""
        : `EUR: ${details.EURDaily/8}`
    }</div><div>${
      details.OtherCurrDaily == "" || details.OtherCurrDaily == null
        ? ""
        : `${details.OtherCurr}: ${details.OtherCurrDaily/8}`
    }</div></td><td>${
      !details.EffectiveDate ? "Not Available" : new Date(details.EffectiveDate).toLocaleDateString()
    }</td><td>${
      !details.FirstName ? "" : details.FirstName
    }</td><td>${
      !details.LastName ? "" : details.LastName
    }</td></tr>`;
  });
  
  drpAss = drpArray.filter (function (value, index, array) { 
    return array.indexOf (value) == index;
  }); 

  for (var i = 0; i < drpAss.length; i++) {
    AssDDHtml += "<option value='" + drpAss[i] + "'>" + drpAss[i] + "</option>";
  }
  
  var htmlForAssitant="<option value=''>Select</option>";
  for (var i = 0; i < arrAssitant.length; i++) 
  {
    htmlForAssitant += "<option value='" + arrAssitant[i] + "'>" + arrAssitant[i] + "</option>";
  }

  ///$("#drpAssistantforEmployee,#drpAssistantforOutside,#drpAssistantforOutside,#drpAssistantforAffiliates,#drpAssistantforAlumni,#drpAssistantforAllPeople").html(AssDDHtml);

  $("#drpAssistantforOutside,#drpAssistantforOutside,#drpAssistantforAffiliates,#drpAssistantforAlumni,#drpAssistantforAllPeople").html(AssDDHtml);

  /* newly added for Assitant*/
  $("#drpAssistantforEmployee,#drpStaffAssitant").html(htmlForAssitant);

  $("#SdhEmpTbody").html(EmpTable);
  $("#SdhOutsideTbody").html(OutTable);
  $("#SdhAffilateTbody").html(AffTable);
  $("#SdhAllumniTbody").html(AlumTable);
  $("#SdhAllPeopleTbody").html(AllDetailsTable);
  $("#SdgofficeinfoTbody").html(OfficeTable);
  $("#SdgBillingrateTbody").html(BillingRateTable);

  var options = {
    language: {
      "emptyTable": "No data available"
    },
    order: [[0, "asc"]],
    lengthMenu: [50, 100],
  };
  bindEmpTable(options);
  bindOutTable(options);
  bindAffTable(options);
  bindAlumTable(options);
  bindAllDetailTable(options);
  bindOfficeTable(options);
  bindBillingRateTable(options);
  SdhEmpTableRowGrouping(1,"StaffAvailabilityTable",bindStaffAvailTable);
  UserProfileDetail();

}
const bindEmpTable = (options) => {
  let EMPTable = (<any>$("#SdhEmpTable")).DataTable(options);

  $("#drpLocationforEmployee").change(function () {
    if ($("#drpLocationforEmployee").val()=="Select") {
      EMPTable.column(4).search("").draw();
    } else {
      EMPTable.column(4).search($("#drpLocationforEmployee option:selected").val()).draw();
    }
  });

  $("#drpTitleforEmployee").change(function () {
    if ($("#drpTitleforEmployee").val()=="Select") {
      EMPTable.column(6).search("").draw();
    } else {
      EMPTable.column(6).search($("#drpTitleforEmployee option:selected").val()).draw();
    }
  });



  $("#drpAssistantforEmployee").change(function () {
    if ($("#drpAssistantforEmployee").val()=="Select") {
      EMPTable.column(7).search("").draw();
    } else {
      EMPTable.column(7).search($("#drpAssistantforEmployee option:selected").val()).draw();
    }
  });

};
const bindOutTable = (options) => {
  let OutTable =(<any>$("#SdhOutsideTable")).DataTable(options);
  $("#drpLocationforOutside").change(function () {
    if ($("#drpLocationforOutside").val()=="Select") {
      OutTable.column(4).search("").draw();
    } else {
      OutTable.column(4).search($("#drpLocationforOutside option:selected").val()).draw();
    }
  });
  $("#drpTitleforOutside").change(function () {
    if ($("#drpTitleforOutside").val()=="Select") {
      OutTable.column(6).search("").draw();
    } else {
      OutTable.column(6).search($("#drpTitleforOutside option:selected").val()).draw();
    }
  });
  $("#drpAssistantforOutside").change(function () {
    if ($("#drpAssistantforOutside").val()=="Select") {
      OutTable.column(7).search("").draw();
    } else {
      OutTable.column(7).search($("#drpAssistantforOutside option:selected").val()).draw();
    }
  });
};
const bindAffTable = (options) => {
  let AffTable=(<any>$("#SdhAffilateTable")).DataTable(options);
  $("#drpLocationforAffiliates").change(function () {
    if ($("#drpLocationforAffiliates").val()=="Select") {
      AffTable.column(4).search("").draw();
    } else {
      AffTable.column(4).search($("#drpLocationforAffiliates option:selected").val()).draw();
    }
  });
  $("#drpTitleforAffiliates").change(function () {
    if ($("#drpTitleforAffiliates").val()=="Select") {
      AffTable.column(6).search("").draw();
    } else {
      AffTable.column(6).search($("#drpTitleforAffiliates option:selected").val()).draw();
    }
  });
  $("#drpAssistantforAffiliates").change(function () {
    if ($("#drpAssistantforAffiliates").val()=="Select") {
      AffTable.column(7).search("").draw();
    } else {
      AffTable.column(7).search($("#drpAssistantforAffiliates option:selected").val()).draw();
    }
  });
};
const bindAlumTable = (options) => {
  let AlumTable= (<any>$("#SdhAllumniTable")).DataTable(options);
  $("#drpLocationforAlumni").change(function () {
    if ($("#drpLocationforAlumni").val()=="Select") {
      AlumTable.column(4).search("").draw();
    } else {
      AlumTable.column(4).search($("#drpLocationforAlumni option:selected").val()).draw();
    }
  });
  $("#drpTitleforAlumni").change(function () {
    if ($("#drpTitleforAlumni").val()=="Select") {
      AlumTable.column(6).search("").draw();
    } else {
      AlumTable.column(6).search($("#drpTitleforAlumni option:selected").val()).draw();
    }
  });
  $("#drpAssistantforAlumni").change(function () {
    if ($("#drpAssistantforAlumni").val()=="Select") {
      AlumTable.column(7).search("").draw();
    } else {
      AlumTable.column(7).search($("#drpAssistantforAlumni option:selected").val()).draw();
    }
  });
};
const bindAllDetailTable = (options) => {
  let AllDetailTable=(<any>$("#SdhAllPeopleTable")).DataTable(options);
  $("#drpLocationforAllPeople").change(function () {
    if ($("#drpLocationforAllPeople").val()=="Select") {
      AllDetailTable.column(4).search("").draw();
    } else {
      AllDetailTable.column(4).search($("#drpLocationforAllPeople option:selected").val()).draw();
    }
  });
  $("#drpTitleforAllPeople").change(function () {
    if ($("#drpTitleforAllPeople").val()=="Select") {
      AllDetailTable.column(6).search("").draw();
    } else {
      AllDetailTable.column(6).search($("#drpTitleforAllPeople option:selected").val()).draw();
    }
  });
  $("#drpAssistantforAllPeople").change(function () {
    if ($("#drpAssistantforAllPeople").val()=="Select") {
      AllDetailTable.column(7).search("").draw();
    } else {
      AllDetailTable.column(7).search($("#drpAssistantforAllPeople option:selected").val()).draw();
    }
  });
};
const bindOfficeTable = (option) => {
  (<any>$("#SdgofficeinfoTable")).DataTable(option);

};
const bindBillingRateTable = (option) => {

 let BillingLoadTable= (<any>$("#SdgBillingrateTable")).DataTable(option);
 $("#drpTitleforBilling").change(function () {
  if ($("#drpTitleforBilling").val()=="Select") {
    BillingLoadTable.column(1).search("").draw();
  } else {
    BillingLoadTable.column(1).search($("#drpTitleforBilling option:selected").val()).draw();
  }
});
};

const bindStaffAvailTable = (option) =>{
let staffAvailTable=  (<any>$('#StaffAvailabilityTable')).DataTable(option);
  $("#drpAffiliatesforBilling").change(function () {
    if ($("#drpAffiliatesforBilling").val()=="Select") {
      staffAvailTable.column(2).search("").draw();
    } else {
      staffAvailTable.column(2).search($("#drpAffiliatesforBilling option:selected").val()).draw();
    }
  });

  $("#drpStaffforBilling").change(function () {
    if ($("#drpStaffforBilling").val()=="Select") {
      staffAvailTable.column(3).search("").draw();
    } else {
      staffAvailTable.column(3).search($("#drpStaffforBilling option:selected").val()).draw();
    }
  });

}

//Todo TableRowGrouping
const SdhEmpTableRowGrouping = (colno, tablename, tablefn) => {
  var collapsedGroups = {};
  var options = {
    order: [[colno, "asc"]],
    language: {
      "emptyTable": "No data available"
    },
    lengthMenu: [50, 100],
    destroy: true,
    rowGroup: {
      dataSrc: colno,
      startRender: function (rows, group) {
        var collapsed = !!collapsedGroups[group];
        rows.nodes().each(function (r) {
          r.style.display = collapsed ? "none" : "";
        });
        return $("<tr/>")
          .append('<td colspan="8">' + group + " (" + rows.count() + ")</td>")
          .attr("data-name", group)
          .toggleClass("collapsed", collapsed);
      },
    },
  };
  $(`#${tablename} tbody`).on("click", "tr.dtrg-start", function () {
    var name = $(this).data("name");
    collapsedGroups[name] = !collapsedGroups[name];
  });
  tablefn(options);
};


function startIt() {
  var schema = {};
  schema["PrincipalAccountType"] = "User,DL,SecGroup,SPGroup";
  schema["SearchPrincipalSource"] = 15;
  schema["ResolvePrincipalSource"] = 15;
  schema["AllowMultipleValues"] = false;
  schema["MaximumEntitySuggestions"] = 50;
  schema["Width"] = "280px";

  SPClientPeoplePicker_InitStandaloneControlWrapper(
    "peoplepickerText",
    null,
    schema
  );
}

const UserProfileDetail = async () => {

  ItemID = 0;
  OfficeAddArr = [];
  SelectedUser = "";
  SelectedUserProfile = [];

  OfficeDetails.forEach((off) => {
    OfficeAddArr.push({ OfficeID: off.ID, OfficePlace: off.Office, OfficeFullAdd: off.Address , Phone: off.Phone});
  });
  
    if(onselecteduser)
    {
      let activeClass = document.querySelectorAll(".nav-items");

      activeClass.forEach((activeC) => {
        activeC["classList"].remove("show");
      });

      $(".clsCollapse").each(function () {
        $(this).removeClass("in").attr("style", "");
      });


      $('#'+onselectedusermain).addClass('show');
      $('#'+onselectedusersec).addClass('in');

      if(onselecteduserTab=="aDirectory")
      {
      $("#availabilityTab").removeClass("active")	
      $("#home").addClass("in active");	
      $("#menu1").removeClass("in active")	
      $("#directoryInfoTab").addClass("active")
      }
      else
      {
        $("#availabilityTab").addClass("active")	
        $("#home").removeClass("in active");	
        $("#menu1").addClass("in active")	
        $("#directoryInfoTab").removeClass("active")
      }
      LoadProfile(onselecteduser);
    }
 
  const username = document.querySelectorAll(".usernametag");
  username.forEach((btn) => {
    btn.addEventListener("click", async (e) => {

      LoadProfile(e.currentTarget["id"]);

    });
  });
  const usernamelast = document.querySelectorAll(".usernametag-last");
  usernamelast.forEach((btn) => {
    btn.addEventListener("click", async (e) => {

      LoadProfile(e.currentTarget["id"]);

    });
  });

      $("#USDDailyEdit").keyup(() => {
        var usdvalue: any = $("#USDDailyEdit").val();
        var finalusdval = usdvalue / 8;
        $("#USDHourlyEdit").val(finalusdval);
      });
      $("#EURDailyEdit").keyup(() => {
        var eurdaily: any = $("#EURDailyEdit").val();
        var finaleurval = eurdaily / 8;
        $("#EURHourlyEdit").val(finaleurval);
      });
      $("#ODailyEdit").keyup(() => {
        var ovalue: any = $("#ODailyEdit").val();
        var finalovalue = ovalue / 8;
        $("#OHourlyEdit").val(finalovalue);
      });
  $(document).on("click", ".clsfileremove", function () {
    let filename = $(this).attr("filename");
    $(this).parent().remove();
    sp.web.getFileByServerRelativeUrl(`/sites/StaffDirectory/BiographyDocument/${SelectedUserProfile[0].Usermail}/${filename}`).recycle().then(function (data) {});
  });
};

const LoadProfile = async(e) =>{
  const sdhEmp = document.querySelector(".sdh-employee");
  const Edit = document.querySelector("#btnEdit");
  const viewBiling = document.querySelector(".view-directory .user-billing-rates");
  const userpage = document.querySelector(".user-profile-page");
  if (!sdhEmp.classList.contains("hide")) {
    sdhEmp.classList.add("hide");
    userpage.classList.remove("hide");
  }

  SelectedUser = e;
 
  SelectedUserProfile = UserDetails.filter((li) => {
    return li.Usermail == SelectedUser;
  });
  selectedUsermail = SelectedUserProfile[0].Usermail;
  $(".profile-picture").attr("src", SelectedUserProfile[0].ProfilePic);

  if((SelectedUserProfile[0].Affiliation != "Employee" && SelectedUserProfile[0].Affiliation != "Outside Consultant")||!SelectedUserProfile[0].showAvailability){
    $("#menu1").addClass("hide");
    $("#availabilityTab").addClass("hide");
  }else{
    $("#menu1").removeClass("hide");
    $("#availabilityTab").removeClass("hide");
  }

  if((SelectedUserProfile[0].Usermail.toLowerCase()==currentMail.toLowerCase()&&IsgeneralStaff)||IsAdminStaff)
  {
     Edit.classList.remove('hide');
     viewBiling.classList.remove('hide');
     $('.BRHead').show();
     if(IsAdminStaff||IssplStaff)
     {
      $('.showAvail-view').show();
      SelectedUserProfile[0].showAvailability? $('#view-show-avail').text("Yes") :$('#view-show-avail').text("No")
     }
     else
     {
      $('.showAvail-view').hide();
     }

  }
  else if(IssplStaff)
  {
    Edit.classList.add('hide');
    viewBiling.classList.remove('hide');
    $('.BRHead').show();
    $('.showAvail-view').show();
    SelectedUserProfile[0].showAvailability? $('#view-show-avail').text("Yes") :$('#view-show-avail').text("No")
  }
  else{
    Edit.classList.add('hide');
    viewBiling.classList.add('hide');
    $('.BRHead').hide();
    $('.showAvail-view').hide();
  }

  if(SelectedUserProfile[0].Affiliation != "Employee" && SelectedUserProfile[0].Affiliation != "Outside Consultant")
  {
    $('.showAvail-view').hide();
  }



  useravailabilityDetails();

            $('#linkedinIDview').html(`<a href="${SelectedUserProfile[0].LinkedInID.Url}" target ='_blank' data-interception="off"><span class="icon-linkedin"></span></a>`);

            $("#user-Designation").html(SelectedUserProfile[0].Affiliation);
            $("#user-staff-function").html(SelectedUserProfile[0].Title);
            $("#user-staff-title").html(SelectedUserProfile[0].JobTitle);

            $("#user-location").html(SelectedUserProfile[0].Location);
            $("#UserProfileName").html(SelectedUserProfile[0].Name);
            $("#UserProfileEmail").html(`<span class="user-mail-icon"></span><a href="mailto:${SelectedUserProfile[0].Usermail}">${SelectedUserProfile[0].Usermail}</a>`);
            if(SelectedUserProfile[0].LinkedInID)
            {
              $("#UserLinkedID").html(`<a href="${SelectedUserProfile[0].LinkedInID.Url}" target ='_blank' data-interception="off"><span class="user-linkedin-icon"></span><p>LinkedIn</p></a>`);
              $("#UserLinkedID").show();
            }
            else
            {
              $("#UserLinkedID").hide();
            }

            if(SelectedUserProfile[0].Skype)
            {
              $("#UserSkypeID").html(`<a href="skype:${SelectedUserProfile[0].Skype}?chat" target ='_blank' data-interception="off"><span class="user-skype-icon"></span><p>Skype</p></a>`);
              $("#UserSkypeID").show();
            }
            else
            {
              $("#UserSkypeID").hide();
            }

            

            //WorkAddress-view-ShowHide

            if(SelectedUserProfile[0].Location)
            {
              $("#WLoctionDetails").html(SelectedUserProfile[0].Location);
              $("#WAddressDetails").html(OfficeAddArr.filter((add) => SelectedUserProfile[0].Location == add.OfficePlace)[0].OfficeFullAdd);
              $('.work-address-view').show();
              $("#officeNoview").parent().removeClass('hide');
              $("#officeNoview").html(OfficeAddArr.filter((add) => SelectedUserProfile[0].Location == add.OfficePlace)[0].Phone);
            }
            else
            {
            $("#officeNoview").parent().addClass('hide')
            $('.work-address-view').hide();       
            $("#WLoctionDetails").html();
            $("#WAddressDetails").html();
            }
            
            
   //personalMail-view-ShowHide

          if(SelectedUserProfile[0].UserPersonalMail) {
            var pemailarr=SelectedUserProfile[0].UserPersonalMail.split(';');
            var pemailHTML="";
            pemailarr.map((email,i)=>{
              if(i!=pemailarr.length-2)
              email? pemailHTML+=`<a href="mailto:${email}">${email}</a><br>`:""
              else
              email? pemailHTML+=`<a href="mailto:${email}">${email}</a>`:""
            });
            $('#userpersonalmail').parent().removeClass('hide');
            $('#userpersonalmail').html(pemailHTML);
          }
          else{
            $('#userpersonalmail').html("")
            $('#userpersonalmail').parent().addClass('hide');
          }

          //Assistant-view-ShowHide
          if(SelectedUserProfile[0].AssisstantName){
            $("#viewAssistant").html(`<h4>Assisstant</h4><div class="d-flex align-item-center">
            <label>Assistant : </label><div class="lblRight" id="assistantViewpage">${SelectedUserProfile[0].AssisstantName}</div>
            </div>`)
          }else{
            $("#viewAssistant").html("")
          }

          //MobileNo-view-ShowHide

          if(SelectedUserProfile[0].PhoneNumber)
          {
            $("#user-phone").show();
            $("#user-phone-l").show();
            var htmlforPhoneNumber="";
            var phnoval=SelectedUserProfile[0].PhoneNumber.split("^");

            if(phnoval.length>0)
            {
              phnoval.map((ph,i)=>{
                if(i!=phnoval.length-2)
                htmlforPhoneNumber+= ph?`${ph}</br>`:``
                else
                htmlforPhoneNumber+=ph?`${ph}`:``
              });
              $("#user-phone").html(htmlforPhoneNumber);
            }
            else
            {
              $("#user-phone").hide();
              $("#user-phone-l").hide();
            }
          }
          else 
          {
            $("#user-phone").hide();
            $("#user-phone-l").hide();
          }

          if(!SelectedUserProfile[0].Location&& !SelectedUserProfile[0].EmergencyNo && !SelectedUserProfile[0].HomeNo&&!SelectedUserProfile[0].SignOther&&!SelectedUserProfile[0].Child )
          {
            $('.contact-info-view').hide();
          }
          else
          {
            //HomeNo-view-ShowHide
            $('.contact-info-view').show();
            if(SelectedUserProfile[0].SignOther) {

              $('#PSignOther').parent().removeClass('hide');
               $('#PSignOther').html(SelectedUserProfile[0].SignOther)
            }
            else{
              $('#PSignOther').html("")
              $('#PSignOther').parent().addClass('hide');
            }

            if(SelectedUserProfile[0].Child) {

              $('#PChildren').parent().removeClass('hide');
               $('#PChildren').html(SelectedUserProfile[0].Child)
            }
            else{
              $('#PChildren').html("")
              $('#PChildren').parent().addClass('hide');
            }

          if(SelectedUserProfile[0].HomeNo)
          {

            $("#homeNoview").parent().removeClass('hide');

          var htmlforHomeNo="";
          var Homenoval=SelectedUserProfile[0].HomeNo.split("^");
          if(Homenoval.length>0)
          {
            Homenoval.map((ho)=>{ho?htmlforHomeNo+=ho+";":""});
            $("#homeNoview").html(htmlforHomeNo);
          }
         else
          {
            $("#homeNoview").parent().addClass('hide')
          }
         }
        else 
        {
          $("#homeNoview").parent().addClass('hide')
        }

       //EmgNo-view-ShowHide

        if(SelectedUserProfile[0].EmergencyNo)
        {
          $("#emergencyNoview").parent().removeClass('hide');
          var htmlforEmergencyNo="";
          var Emnoval=SelectedUserProfile[0].EmergencyNo.split("^");
          if(Emnoval.length>0)
          {
            Emnoval.map((EO)=>{EO?htmlforEmergencyNo+=EO+";":""});
            $("#emergencyNoview").html(htmlforEmergencyNo);
          }
          else{
            $("#emergencyNoview").parent().addClass('hide');
          }
        }
        else
        {
          $("#emergencyNoview").parent().addClass('hide');
        }

        //OfficeNo-view-ShowHide

        // if(SelectedUserProfile[0].OfficeNo)
        // {
        //   $("#officeNoview").parent().removeClass('hide')
        //   var htmlforOfficeNo="";
        //   var Ofnoval=SelectedUserProfile[0].OfficeNo.split("^");
        //   if(Ofnoval.length>0)
        //   {
        //     Ofnoval.map((OO)=>{OO?htmlforOfficeNo+=OO+";":""});
        //     $("#officeNoview").html(htmlforOfficeNo);
        //   }
        //   else{
        //     $("#officeNoview").parent().addClass('hide')
        //   }
        // }
        // else 
        // {
        //   $("#officeNoview").parent().addClass('hide')
        // }

    }

          if(!SelectedUserProfile[0].HAddLine && !SelectedUserProfile[0].HAddCity && !SelectedUserProfile[0].HAddState&&! SelectedUserProfile[0].HAddPCode && !SelectedUserProfile[0].HAddPCountry && !SelectedUserProfile[0].SignOther && !SelectedUserProfile[0].Child)
          {
            $('.personal-info-view').hide();
          }
          else
          {
            $('.personal-info-view').show();
            $('#PAddLine').parent().removeClass('hide');
            $('#PAddCity').parent().addClass('hide');
            var addressHTML="";
            SelectedUserProfile[0].HAddLine? addressHTML+=`${SelectedUserProfile[0].HAddLine}<br>`:""
            SelectedUserProfile[0].HAddCity? addressHTML+=`${SelectedUserProfile[0].HAddCity}<br>`:""
            SelectedUserProfile[0].HAddState? addressHTML+=`${SelectedUserProfile[0].HAddState}<br>`:""
            SelectedUserProfile[0].HAddPCode? addressHTML+=`${SelectedUserProfile[0].HAddPCode}<br>`:""
            SelectedUserProfile[0].HAddPCountry? addressHTML+=`${SelectedUserProfile[0].HAddPCountry}<br>`:""

            $('#PAddLine').html(addressHTML);

          }

          // if(!SelectedUserProfile[0].HAddLine && !SelectedUserProfile[0].HAddCity && !SelectedUserProfile[0].HAddState&&! SelectedUserProfile[0].HAddPCode && !SelectedUserProfile[0].HAddPCountry && !SelectedUserProfile[0].SignOther && !SelectedUserProfile[0].Child)
          // {
          //   $('.personal-info-view').hide();
          // }
          // else{
          //   $('.personal-info-view').show();
          //   if(SelectedUserProfile[0].HAddLine) {

          //     $('#PAddLine').parent().removeClass('hide');
          //      $('#PAddLine').html(SelectedUserProfile[0].HAddLine)
          //   }
          //   else{
          //     $('#PAddLine').html("")
          //     $('#PAddLine').parent().addClass('hide');
          //   }
          //   if(SelectedUserProfile[0].HAddCity) {

          //     $('#PAddCity').parent().removeClass('hide');
          //      $('#PAddCity').html(SelectedUserProfile[0].HAddCity)
          //   }
          //   else{
          //     $('#PAddCity').html("")
          //     $('#PAddCity').parent().addClass('hide');
          //   }
          //   if(SelectedUserProfile[0].HAddState) {

          //     $('#PAddState').parent().removeClass('hide');
          //      $('#PAddState').html(SelectedUserProfile[0].HAddState)
          //   }
          //   else{
          //     $('#PAddState').html("")
          //     $('#PAddState').parent().addClass('hide');
          //   }
          //   if(SelectedUserProfile[0].HAddPCode) {

          //     $('#PAddPCode').parent().removeClass('hide');
          //      $('#PAddPCode').html(SelectedUserProfile[0].HAddPCode)
          //   }
          //   else{
          //     $('#PAddPCode').html("")
          //     $('#PAddPCode').parent().addClass('hide');
          //   }
          //   if(SelectedUserProfile[0].HAddPCountry) {

          //     $('#PAddPCountry').parent().removeClass('hide');
          //      $('#PAddPCountry').html(SelectedUserProfile[0].HAddPCountry)
          //   }
          //   else{
          //     $('#PAddPCountry').html("")
          //     $('#PAddPCountry').parent().addClass('hide');
          //   }
          // }

          //StaffStatus-view-ShowHide
          if(SelectedUserProfile[0].StaffStatus)
          {
            $("#staffStatus").html(SelectedUserProfile[0].StaffStatus);
            $("#workscheduleViewSec").html(SelectedUserProfile[0].StaffStatus == "Part-time"? `<div class="d-flex"><label>Work Schedule</label><p class="lblRight" id="workSchedule">${SelectedUserProfile[0].WorkSchedule ?SelectedUserProfile[0].WorkSchedule:""}</p></div>` : "");
            $('.StaffStatus-view').show();
          }
          else
          {
            $('.StaffStatus-view').hide();
          }

          //Citizen-view-ShowHide

          if(SelectedUserProfile[0].Citizen) {

            $('#citizenship').parent().removeClass('hide');
             $('#citizenship').html(SelectedUserProfile[0].Citizen)
          }
          else{
            $('#citizenship').html("")
            $('#citizenship').parent().addClass('hide');
          }

   //BillingRate -view-ShowHide

  let billingRateHtml = "";
  if(!SelectedUserProfile[0].USDDaily&&!SelectedUserProfile[0].EURDaily&&!SelectedUserProfile[0].OtherCurrDaily)
  {
    viewBiling.classList.add('hide');
    $('.BRHead').hide();
  }
  else if(IssplStaff||IsAdminStaff||(SelectedUserProfile[0].Usermail.toLowerCase()==currentMail.toLowerCase()&&IsgeneralStaff))
  {
    viewBiling.classList.remove('hide');
    $('.BRHead').show();
    if (SelectedUserProfile[0].USDDaily != null && SelectedUserProfile[0].USDDaily != 0 && SelectedUserProfile[0].USDDaily != "0")
    {
var USDHourly=SelectedUserProfile[0].USDDaily/8;
      billingRateHtml += `<div class="billing-rates"><label>USD Daily Rate</label><div class="usd-daily-rate lblBlue" id="UsdDailyRate">${SelectedUserProfile[0].USDDaily}</div></div><div class="billing-rates"><label>USD Hourly Rate</label><div class="usd-hourly-rate lblBlue" id="UsdHourlyRate">${USDHourly}</div></div>`;
    }

    if (SelectedUserProfile[0].EURDaily != null &&SelectedUserProfile[0].EURDaily != 0 &&SelectedUserProfile[0].EURDaily != "0") 
    {
      var EURHourly=SelectedUserProfile[0].EURDaily/8;
      billingRateHtml += `<div class="billing-rates"><label>EUR Daily Rate</label><div class="eur-daily-rate lblBlue" id="EURDailyRate">${SelectedUserProfile[0].EURDaily}</div></div><div class="billing-rates"><label>EUR Hourly Rate</label><div class="eur-hourly-rate lblBlue" id="EURHourlyRate">${EURHourly}</div></div>`;
    }

    if ( SelectedUserProfile[0].OtherCurrDaily != null &&SelectedUserProfile[0].OtherCurrDaily != 0 &&SelectedUserProfile[0].OtherCurrDaily != "0") 
    {
      var OtherCurrHourly=SelectedUserProfile[0].OtherCurrDaily/8;
      billingRateHtml += `<div class="billing-rates"><label>${SelectedUserProfile[0].OtherCurr} Daily Rate</label><div class="eur-daily-rate lblBlue" id="oDailyRate">${SelectedUserProfile[0].OtherCurrDaily}</div></div><div class="billing-rates"><label>${SelectedUserProfile[0].OtherCurr} Hourly Rate</label><div class="eur-hourly-rate lblBlue" id="oHourlyRate">${OtherCurrHourly}</div></div>`;
    }
    if (SelectedUserProfile[0].EffectiveDate != null) 
    {
      billingRateHtml += ` <div class="billing-effective-date"><label>Effective Date</label><div class="effective-date lblBlue" id="EffectiveDate">${new Date(
        SelectedUserProfile[0].EffectiveDate
      ).toLocaleDateString()}</div></div>`;
    }
    if (SelectedUserProfile[0].BillingRateComments != null) 
    {
      billingRateHtml += ` <div class="billing-rates"><label>Comments</label><div class="Billing-comments" id="BillingRateComments">${SelectedUserProfile[0].BillingRateComments}</div></div>`;
    }   
    
    $("#BillingRateDetails").html(billingRateHtml);

  }

  //ShortBio-view-ShowHide

    if(SelectedUserProfile[0].ShortBio) {
      $('#shortbioh').show();
      $('#shortbio').html(SelectedUserProfile[0].ShortBio)
    }
    else{
      $('#shortbio').hide();
      $('#shortbioh').hide();
    }

   //bioAttachment-view-ShowHide
  
  let filesHtml = "";
  let editfileHtml = "";
  let files = await sp.web.getFolderByServerRelativeUrl(`BiographyDocument/${SelectedUserProfile[0].Usermail}`).files.get();
  if(files.length>0)
  {     
     files.forEach((file) => {
    if ( file.Name.split(".").pop() == "doc" ||file.Name.split(".").pop() == "docx") 
    {
      filesHtml += `<div class="doc-section"><span class="word-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank" data-interception="off">${file.Name}</a></div>`;
      editfileHtml += `<div class="quantityFiles"><span class="upload-filename">${file.Name}</span><a filename="${file.Name}" class="clsfileremove">x</a></div>`;
    }
   else if (file.Name.split(".").pop() == "xlsx" || file.Name.split(".").pop() == "csv") {
      filesHtml += `<div class="doc-section"><span class="excel-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank" data-interception="off">${file.Name}</a></div>`;
      editfileHtml += `<div class="quantityFiles"><span class="upload-filename">${file.Name}</span><a  filename="${file.Name}" class="clsfileremove">x</a></div>`;
    } 
    else if (file.Name.split(".").pop() == "png" ||file.Name.split(".").pop() == "jpg" ||file.Name.split(".").pop() == "jpeg")
    {
      filesHtml += `<div class="doc-section"><span class="pic-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank" data-interception="off">${file.Name}</a></div>`;
      editfileHtml += `<div class="quantityFiles"><span class="upload-filename">${file.Name}</span><a  filename="${file.Name}" class="clsfileremove">x</a></div>`;
    } 
    else {
      filesHtml += `<div class="doc-section"><span class="new-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank" data-interception="off">${file.Name}</a></div>`;
      editfileHtml += `<div class="quantityFiles"><span class="upload-filename">${file.Name}</span><a  filename="${file.Name}" class="clsfileremove">x</a></div>`;
    }
    $("#bioAttachment").html(filesHtml);
    $("#filesfromfolder").html(editfileHtml);
    $('#bioAttachhead').show();
  });

  }
  else
  {
    $('#bioAttachhead').hide();
    $("#bioAttachment").html("");
    $("#filesfromfolder").html("");
  }



    ItemID = SelectedUserProfile[0].ItemID;

    //Experience-view-ShowHide

    if(!SelectedUserProfile[0].SpecialKnowledge&&! SelectedUserProfile[0].Membership&& !SelectedUserProfile[0].Software &&! SelectedUserProfile[0].SDGCourse && !SelectedUserProfile[0].Language  && !SelectedUserProfile[0].Industry)
    {
       $('.other-exp-view').hide();
    }
    else
    {
      $('.other-exp-view').show();
      if(SelectedUserProfile[0].Industry) 
      {
        $("#industryExpViewSec").removeClass("hide")	
        // $('#IndustryExpl').show();	
        $('#IndustryExp').html(SelectedUserProfile[0].Industry)
      }
      else{
	      //  $('#IndustryExp').hide();	
      //  $('#IndustryExpl').hide();	
      $("#industryExpViewSec").addClass("hide")
      }

      if(SelectedUserProfile[0].Language) 
      {
        $("#languageExpViewSec").removeClass("hide");	
        // $('#LanguageExpl').show();	
      $('#LanguageExp').html(SelectedUserProfile[0].Language)
      }
      else
      {
        $("#languageExpViewSec").addClass("hide");	
        // $('#LanguageExp').hide();	
        // $('#LanguageExpl').hide();
      }

      if(SelectedUserProfile[0].SDGCourse)
      {
        $("#courseExpViewSec").removeClass("hide");	
        // $('#SDGCoursel').show()	
        $('#SDGCourse').html(SelectedUserProfile[0].SDGCourse);
      }
     else
      {
        $("#courseExpViewSec").addClass("hide");	
        // $('#SDGCourse').hide();	
        // $('#SDGCoursel').hide();
      }

      if(SelectedUserProfile[0].Software)
      {
        $("#softwareExpViewSec").removeClass("hide")	
        // $('#SoftwareExpl').show();	
      $('#SoftwareExp').html(SelectedUserProfile[0].Software)
      }
      else
      {
        $("#softwareExpViewSec").addClass("hide")	
        // $('#SoftwareExp').hide();	
        // $('#SoftwareExpl').hide();
      }

      if(SelectedUserProfile[0].Membership) 
      {
        $("#membershipViewSec").removeClass("hide")	
        // $('#MembershipExpl').show();	
        $('#MembershipExp').html(SelectedUserProfile[0].Membership)
      }
      else
      {
        $("#membershipViewSec").addClass("hide")	
        // $('#MembershipExp').hide();	
        // $('#MembershipExpl').hide();
      } 

      if(SelectedUserProfile[0].SpecialKnowledge) 
      {
        $("#specialKnowledgeViewSec").removeClass("hide");	
        // $('#SpecialKnowledgel').show();	
        $('#SpecialKnowledge').html(SelectedUserProfile[0].SpecialKnowledge)
      }
      else
      {
        $("#specialKnowledgeViewSec").addClass("hide");	
        // $('#SpecialKnowledge').hide();	
        // $('#SpecialKnowledgel').hide();
      } 

    }

  if($('#shortbioh').css('display') == 'none'&&$('#bioAttachhead').css('display') == 'none'&&$('.other-exp-view').css('display') == 'none')
  {
    $('.Biography-Experience-view').hide()
  }
  else
  {
    $('.Biography-Experience-view').show()
  }

    var editfMonth: any = "";
    var editfday: any = "";
    var Sdd = new Date(SelectedUserProfile[0].EffectiveDate).getDate();
    Sdd < 10 ? (editfday = "0" + Sdd) : (editfday = Sdd);
    var Smm = new Date(SelectedUserProfile[0].EffectiveDate).getMonth() + 1;
    Smm < 10 ? (editfMonth = "0" + Smm) : (editfMonth = Smm);
    var Syyyy = new Date(SelectedUserProfile[0].EffectiveDate).getFullYear();
    var Sdateformat = Syyyy + "-" + editfMonth + "-" + editfday;

    $("#EffectiveDateEdit").val(Sdateformat);

}

const editFunction = async() => {

  let LocOptionHtml = "";
  let listLocation  = await sp.web.getList(listUrl + "SDGOfficeInfo").items.get(); 
  let withoutAlumini = listLocation.filter((li)=>li.AlumniOffice == false);
// Check Alumni for Dropdown
  if(SelectedUserProfile[0].Affiliation == "Alumni"){
  
  listLocation.forEach((li) => {
    LocOptionHtml += `<option value="${li.ID}">${li.Office}</option>`;
  }); 
}
  else{
    withoutAlumini.forEach((li) => {
      LocOptionHtml += `<option value="${li.ID}">${li.Office}</option>`;
    });
  }
    $('#workLocationDD').html("");
    $('#workLocationDD').html(LocOptionHtml);
  await SPComponentLoader.loadScript("/_layouts/15/init.js").then(() => {});
  await SPComponentLoader.loadScript("/_layouts/15/MicrosoftAjax.js").then(() => {});
  await SPComponentLoader.loadScript("/_layouts/15/1033/sts_strings.js");
  await SPComponentLoader.loadScript("/_layouts/15/clientforms.js");
  await SPComponentLoader.loadScript("/_layouts/15/clienttemplates.js");
  await SPComponentLoader.loadScript("/_layouts/15/clientpeoplepicker.js");
  await SPComponentLoader.loadScript("/_layouts/15/autofill.js");
  await SPComponentLoader.loadScript("/_layouts/15/sp.runtime.js");
  await SPComponentLoader.loadScript("/_layouts/15/SP.js");
  await SPComponentLoader.loadScript("/_layouts/15/SP.js");
  await SPComponentLoader.loadScript("/_layouts/15/SP.js");
  await SPComponentLoader.loadScript("/_layouts/15/sp.core.js");


  await startIt();
  const Edit = document.querySelector("#btnEdit");
  const UserView = document.querySelector(".view-directory");
  const UserEdit = document.querySelector(".edit-directory");
  const adminviewBilling = document.querySelector("#BillingRateDetailsEdit");
  const userviewBilling = document.querySelector("#BillingRateDetailsView");
  const viewBiling =  document.querySelector(".edit-directory .user-billing-rates");
 
  if (!UserView.classList.contains("hide")) {
    UserView.classList.add("hide");
    UserEdit.classList.remove("hide"); 
    Edit.classList.add("hide");
    if(IsAdminStaff)
    {
      viewBiling.classList.remove('hide');
      adminviewBilling.classList.remove('hide');
      userviewBilling.classList.add('hide');
      $('.BRHead').show();
      $('.show-availablity-sec').show()
    }

    else
    {
      viewBiling.classList.add('hide');
      adminviewBilling.classList.add('hide');
      userviewBilling.classList.remove('hide');
      $('.BRHead').hide();
      $('.show-availablity-sec').hide()

     if(SelectedUserProfile[0].Affiliation != "Employee" && SelectedUserProfile[0].Affiliation != "Outside Consultant")
     {
      $('.show-availablity-sec').hide()
     }


      let billingRateHtml = "";
      if (SelectedUserProfile[0].USDDaily != null && SelectedUserProfile[0].USDDaily != 0 && SelectedUserProfile[0].USDDaily != "0")
      {

        billingRateHtml += `<div class="billing-rates"><label>USD Daily Rate</label><div class="usd-daily-rate lblBlue" id="UsdDailyRate">${SelectedUserProfile[0].USDDaily}</div></div><div class="billing-rates"><label>USD Hourly Rate</label><div class="usd-hourly-rate lblBlue" id="UsdHourlyRate">${SelectedUserProfile[0].USDDaily/8}</div></div>`;
      }

      if (
        SelectedUserProfile[0].EURDaily != null &&
        SelectedUserProfile[0].EURDaily != 0 &&
        SelectedUserProfile[0].EURDaily != "0"
      ) {
        billingRateHtml += `<div class="billing-rates"><label>EUR Daily Rate</label><div class="eur-daily-rate lblBlue" id="EURDailyRate">${SelectedUserProfile[0].EURDaily}</div></div><div class="billing-rates"><label>EUR Hourly Rate</label><div class="eur-hourly-rate lblBlue" id="EURHourlyRate">${SelectedUserProfile[0].EURDaily/8}</div></div>`;
      }

      if (
        SelectedUserProfile[0].OtherCurrDaily != null &&
        SelectedUserProfile[0].OtherCurrDaily != 0 &&
        SelectedUserProfile[0].OtherCurrDaily != "0"
      ) {
        billingRateHtml += `<div class="billing-rates"><label>${SelectedUserProfile[0].OtherCurr} Daily Rate</label><div class="eur-daily-rate lblBlue" id="oDailyRate">${SelectedUserProfile[0].OtherCurrDaily}</div></div><div class="billing-rates"><label>${SelectedUserProfile[0].OtherCurr} Hourly Rate</label><div class="eur-hourly-rate lblBlue" id="oHourlyRate">${SelectedUserProfile[0].OtherCurrDaily/8}</div></div>`;
      }
      if (SelectedUserProfile[0].EffectiveDate != null) {
        billingRateHtml += ` <div class="billing-effective-date"><label>Effective Date</label><div class="effective-date lblBlue" id="EffectiveDate">${new Date(
          SelectedUserProfile[0].EffectiveDate
        ).toLocaleDateString()}</div></div>`;
      }
      if (
        SelectedUserProfile[0].BillingRateComments != null
      ) {
        billingRateHtml += ` <div class="billing-rates"><label>Comments</label><div class="Billing-comments" id="BillingRateComments">${SelectedUserProfile[0].BillingRateComments}</div></div>`;}
      $('#BillingRateDetailsView').html("");
      $('#BillingRateDetailsView').html(billingRateHtml)
    }

    if(SelectedUserProfile[0].Affiliation != "Employee" && SelectedUserProfile[0].Affiliation != "Outside Consultant")
    {
     $('.show-availablity-sec').hide()
    }


  } else {
    UserEdit.classList.remove("hide");
    Edit.classList.add("hide");
  }

  let MobileNumberHtmlSec = "";
  let HomeNumberHtmlSec = "";
  let EmergencyNumberHtmlSec = "";
  let OfficeNumberHtmlSec = "";

  SelectedUserProfile[0].showAvailability?$('#show-availability').prop('checked',true):$('#show-availability').prop('checked',false)

  $("#EditedAddressDetails").html(OfficeAddArr.filter((add) => SelectedUserProfile[0].Location == add.OfficePlace)[0].OfficeFullAdd);

  let MCCodeArr = []
  if (
    SelectedUserProfile[0].PhoneNumber != "" && SelectedUserProfile[0].PhoneNumber != null
  ) {
    let AllMnumber = SelectedUserProfile[0].PhoneNumber.split("^");
    AllMnumber.pop();
    let AllMobileNumbers = AllMnumber;

    AllMobileNumbers.forEach((numbers, i) => {
      // let SplitedMNum = numbers.split(" - ");
      // MCCodeArr.push(SplitedMNum[0])

      if (i == 0) {
        MobileNumberHtmlSec += `<div class="d-flex mobNumbers"><input type="text" class="mobNo" id="" value="${numbers}"><span class="addMobNo add-icon"></div>`;
        $("#mobileNoSec").html(MobileNumberHtmlSec);
      } else {
        MobileNumberHtmlSec += `<div class="d-flex mobNumbers"><input type="text" class="mobNo" id="" value="${numbers}"><span class="removeMobNo remove-icon"></div>`;
        $("#mobileNoSec").html(MobileNumberHtmlSec);
      }
      
    });

  } else {
    MobileNumberHtmlSec += `<div class="d-flex mobNumbers"><input type="text" class="mobNo" id=""><span class="addMobNo add-icon"></div>`;
    $("#mobileNoSec").html(MobileNumberHtmlSec);
  }
  let HCCodeArr=[];
if(SelectedUserProfile[0].HomeNo != "" && SelectedUserProfile[0].HomeNo != null){
  let AllHNumber = SelectedUserProfile[0].HomeNo.split("^");
  AllHNumber.pop();
  let AllHomeNumber = AllHNumber;
  AllHomeNumber.forEach((hnumbs,j)=>{
    // let SplitedHNum = hnumbs.split(' - ');
    // HCCodeArr.push(SplitedHNum[0])
    if(j == 0){
      HomeNumberHtmlSec +=`<div class="d-flex homeNumbers"><input type="text" class="home" id="" value="${hnumbs}"><span class="addHomeNo add-icon"></div>`
      $('#homeNoSec').html(HomeNumberHtmlSec);
    }else{
      HomeNumberHtmlSec +=`<div class="d-flex homeNumbers"><input type="text" class="home" id="" value="${hnumbs}"><span class="removeHomeNo remove-icon"></div>`
      $('#homeNoSec').html(HomeNumberHtmlSec);
    }
  });
 
}
else{
  HomeNumberHtmlSec +=`<div class="d-flex homeNumbers"><input type="text" class="home" id=""><span class="addHomeNo add-icon"></div>`
  $('#homeNoSec').html(HomeNumberHtmlSec);
}

let ECCodeArr = [];
if(SelectedUserProfile[0].EmergencyNo != ""  && SelectedUserProfile[0].EmergencyNo != null){
  let AllENumber = SelectedUserProfile[0].EmergencyNo.split("^");
  AllENumber.pop();
  let AllEmergencyNumber = AllENumber;
  AllEmergencyNumber.forEach((enums,k)=>{
    // let SplitedENum = enums.split(' - ');
    // ECCodeArr.push(SplitedENum[0])
    if(k==0){
      EmergencyNumberHtmlSec +=`<div class="d-flex emergencyNumbers"><input type="text" class="home" id="" value="${enums}"><span class="addEmergencyNo add-icon"></div>`
      $('#emergencyNoSec').html(EmergencyNumberHtmlSec);
    }else{
      EmergencyNumberHtmlSec +=`<div class="d-flex emergencyNumbers"><input type="text" class="home" id="" value="${enums}"><span class="removeEmergencyNo remove-icon"></div>`
      $('#emergencyNoSec').html(EmergencyNumberHtmlSec);
    }
  })
  
}else{
  EmergencyNumberHtmlSec +=`<div class="d-flex emergencyNumbers"><input type="text" class="home" id=""><span class="addEmergencyNo add-icon"></div>`
  $('#emergencyNoSec').html(EmergencyNumberHtmlSec);
}

let OCCodeArr = [];
if(SelectedUserProfile[0].OfficeNo != ""  && SelectedUserProfile[0].OfficeNo != null){
  let AllONumber = SelectedUserProfile[0].OfficeNo.split("^");
  AllONumber.pop();
  let AllOfficeNumber = AllONumber;
  AllOfficeNumber.forEach((onums,l)=>{
    // let SplitedONum = onums.split(' - ');
    // OCCodeArr.push(SplitedONum[0])
    if(l==0){
      OfficeNumberHtmlSec +=`<div class="d-flex officeNumbers"><input type="text" class="home" id="" value="${onums}"><span class="addOfficeNo add-icon"></div>`
      $('#officeNoSec').html(OfficeNumberHtmlSec);
    }else{
      OfficeNumberHtmlSec +=`<div class="d-flex officeNumbers"><input type="text" class="home" id="" value="${onums}"><span class="removeOfficeNo remove-icon"></div>`
      $('#officeNoSec').html(OfficeNumberHtmlSec);
    }
  })
}else{
  OfficeNumberHtmlSec +=`<div class="d-flex officeNumbers"><select class="officeNoCode">${CCodeHtml}</select><input type="text" class="home" id=""><span class="addOfficeNo add-icon"></div>`
  $('#officeNoSec').html(OfficeNumberHtmlSec);
}

if(OCCodeArr.length > 0){
  $('.officeNoCode').each((i,evt)=>{
    var ecID=OCCodeArr[i]
    var idx=CCodeArr.indexOf(ecID)
    evt["selectedIndex"]=idx
    // $(this).value=ecID
    // $(this).val(ECCodeArr[i])
    // $("#"+evt.id).val(ECCodeArr[i])
  })
}

if(ECCodeArr.length > 0){
  $('.emergencyNoCode').each((i,evt)=>{
    var ecID=ECCodeArr[i]
    var idx=CCodeArr.indexOf(ecID)
    evt["selectedIndex"]=idx
    // $(this).value=ecID
    // $(this).val(ECCodeArr[i])
    // $("#"+evt.id).val(ECCodeArr[i])
  })
}


if(MCCodeArr.length > 0){
  $('.mobNoCode').each((i,evt)=>{
    var ecID=MCCodeArr[i]
    var idx=CCodeArr.indexOf(ecID)
    evt["selectedIndex"]=idx
    // $(this).value=ecID
    // $(this).val(ECCodeArr[i])
    // $("#"+evt.id).val(ECCodeArr[i])
  })
}


if(HCCodeArr.length > 0){
  $('.homeNoCode').each((i,evt)=>{
    var ecID=HCCodeArr[i]
    var idx=CCodeArr.indexOf(ecID)
    evt["selectedIndex"]=idx
    // $(this).value=ecID
    // $(this).val(ECCodeArr[i])
    // $("#"+evt.id).val(ECCodeArr[i])
  })
}

  $(".addMobNo").click(() => {
    multipleMobNo();
  });
  $(".addHomeNo").click(() => {
    multipleHomeNo();
  });
  $(".addEmergencyNo").click(() => {
    multipleEmergencyNo();
  });
  $(".addOfficeNo").click(() => {
    multipleOfficeNo();
  });
var ofcAdd=[]
  if(SelectedUserProfile[0].Location)
  {
    ofcAdd= OfficeAddArr.filter(
      (add) => SelectedUserProfile[0].Location == add.OfficePlace
    );
  }
  else
  {
    $("#EditedAddressDetails").html("");

  }

  var reg1 = new RegExp("<div class=\"ExternalClass[0-9A-F]+\">", "");  
  var reg2 = new RegExp("</div>$", ""); 
  var fetchInVal=SelectedUserProfile[0].Industry.replace(reg1, "").replace(reg2, "");
  fetchInVal=fetchInVal.split("<br>").join("\n");



  var fetchLanguageVal=SelectedUserProfile[0].Language.replace(reg1, "").replace(reg2, "");
  fetchLanguageVal=fetchLanguageVal.split("<br>").join("\n");



  var fetchSDGCourseVal=SelectedUserProfile[0].SDGCourse.replace(reg1, "").replace(reg2, "");
  fetchSDGCourseVal=fetchSDGCourseVal.split("<br>").join("\n");



  var fetchSoftwareVal=SelectedUserProfile[0].Software.replace(reg1, "").replace(reg2, "");
  fetchSoftwareVal=fetchSoftwareVal.split("<br>").join("\n");



  var fetchMembershipVal=SelectedUserProfile[0].Membership.replace(reg1, "").replace(reg2, "");
  fetchMembershipVal=fetchMembershipVal.split("<br>").join("\n");


  var fetchSpecialKnowledgeVal=SelectedUserProfile[0].SpecialKnowledge.replace(reg1, "").replace(reg2, "");
  fetchSpecialKnowledgeVal=fetchSpecialKnowledgeVal.split("<br>").join("\n");



  var fetchShortBioVal=SelectedUserProfile[0].ShortBio.replace(reg1, "").replace(reg2, "");
  fetchShortBioVal=fetchShortBioVal.split("<br>").join("\n");
console.log(SelectedUserProfile[0].Location);

  $("#StaffFunctionEdit").val(SelectedUserProfile[0].Title);
  $("#StaffAffiliatesEdit").val(SelectedUserProfile[0].Affiliation);
  $("#PAddLineE").val(SelectedUserProfile[0].HAddLine);
  $("#PAddCityE").val(SelectedUserProfile[0].HAddCity);
  $("#PAddStateE").val(SelectedUserProfile[0].HAddState);
  $("#PAddPCodeE").val(SelectedUserProfile[0].HAddPCode);
  $("#PAddCountryE").val(SelectedUserProfile[0].HAddPCountry);
  $("#Eshortbio").val(fetchShortBioVal);
  $("#EIndustry").val(fetchInVal);
  $("#ELanguage").val(fetchLanguageVal);
  $("#ESDGCourse").val(fetchSDGCourseVal);
  $("#ESoftwarExp").val(fetchSoftwareVal);
  $("#EMembership").val(fetchMembershipVal);
  $("#ESKnowledge").val(fetchSpecialKnowledgeVal);
  $("#citizenshipE").val(SelectedUserProfile[0].Citizen);
  $("#linkedInID").val(SelectedUserProfile[0].LinkedInID.Url);
  $("#SkypeID").val(SelectedUserProfile[0].Skype);
  $("#children").val(SelectedUserProfile[0].Child);
  $("#significantOther").val(SelectedUserProfile[0].SignOther);
  $("#USDDailyEdit").val(SelectedUserProfile[0].USDDaily);
  $("#USDHourlyEdit").val(SelectedUserProfile[0].USDDaily/8);
  $("#EURDailyEdit").val(SelectedUserProfile[0].EURDaily);
  $("#EURHourlyEdit").val(SelectedUserProfile[0].EURDaily/8);
  $("#personalmailID").val(SelectedUserProfile[0].UserPersonalMail);
  $("#workLocationDD").val(SelectedUserProfile[0].Locationid);
  $("#staffstatusDD").val(SelectedUserProfile[0].StaffStatus);
  $("#othercurrDD").val(SelectedUserProfile[0].OtherCurr);
  $("#ODailyEdit").val(SelectedUserProfile[0].OtherCurrDaily);
  $("#OHourlyEdit").val(SelectedUserProfile[0].OtherCurrDaily/8);
  $("#BillingRateCommentsEdit").val(SelectedUserProfile[0].BillingRateComments);

  if(SelectedUserProfile[0].AssisstantName)
  $("#drpStaffAssitant").val(SelectedUserProfile[0].AssisstantName);
  else
  $("#drpStaffAssitant").val("");

  if(SelectedUserProfile[0].StaffStatus == "Part-time"){
    $("#workscheduleEdit").html("");
    $("#workscheduleEdit").html(`<div class="d-flex w-100" id="workscheduleSec"> <label>Work Schedule</label><div class="w-100"><input type="text" id="workScheduleE" value="${SelectedUserProfile[0].WorkSchedule == "" || SelectedUserProfile[0].WorkSchedule == null ? "":SelectedUserProfile[0].WorkSchedule}"></div></div>`)
  }
  else{
    $("#workscheduleEdit").html("");
    $("#workscheduleEdit").html(`<div class="d-flex w-100 hide" id="workscheduleSec"> <label>Work Schedule</label><div class="w-100"><input type="text" id="workScheduleE" value="${SelectedUserProfile[0].WorkSchedule == "" || SelectedUserProfile[0].WorkSchedule == null ? "":SelectedUserProfile[0].WorkSchedule}"></div></div>`)
  }
  if(SelectedUserProfile[0].AssistantMail)
  {
    var emailAddress =
    "i:0#.f|membership|" + SelectedUserProfile[0].AssistantMail.toLowerCase();
  var divID = "peoplepickerText_TopSpan";
  SPClientPeoplePicker.SPClientPeoplePickerDict[divID].AddUnresolvedUser(
    {
      Key: emailAddress,
      DisplayText: SelectedUserProfile[0].Assistant,
      Email: SelectedUserProfile[0].AssistantMail.toLowerCase(),
    },
    true
  );
  }
};
const editsubmitFunction = async () => {
  let mobNumUpdate = "";
  let homeNumUpdate = "";
  let emergencyNumUpdate = "";
  let officeNumUpdate = "";
  let mobNumbers = document.querySelectorAll(".mobNumbers");
  let homeNumbers = document.querySelectorAll(".homeNumbers");
  let emergencyNumbers = document.querySelectorAll(".emergencyNumbers");
  let officeNumbers = document.querySelectorAll(".officeNumbers");
  mobNumbers.forEach((nums) => {
    if(nums.children[0]["value"] != ""){
      mobNumUpdate += `${
        nums.children[0]["value"]
      }^`;
    }
  });
  homeNumbers.forEach((nums) => {
    if(nums.children[0]["value"] != ""){
    homeNumUpdate += `${nums.children[0]["value"]}^`;
  }
  });
  emergencyNumbers.forEach((nums) => {
    if(nums.children[0]["value"] != ""){
    emergencyNumUpdate += `${nums.children[0]["value"]}^`;
  }
  });
  officeNumbers.forEach((nums) => {
    if(nums.children[0]["value"] != ""){
    officeNumUpdate += `${nums.children[0]["value"]}^`;
  }
  });

  if (bioAttachArr.length > 0) {
    bioAttachArr.map((filedata) => {
      sp.web.folders
        .add(`/sites/StaffDirectory/BiographyDocument/${selectedUsermail}`)
        .then((data) => {
          sp.web
            .getFolderByServerRelativeUrl(data.data.ServerRelativeUrl)
            .files.add(filedata.name, filedata, true);
        });
    });
  }
  var dispTitle = "APickerField";
  var pickerDiv = $("[id$='peoplepickerText'][title='" + dispTitle + "']");
  var peoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict;
  var userInfo = peoplePicker.peoplepickerText_TopSpan.GetAllUserInfo();
  let profileID = 0
  if(userInfo.length >0){
    const loginName = userInfo[0].Key.split("|")[2];
    const profile = await sp.web.siteUsers.getByEmail(loginName).get();
    profileID = profile.Id
  }
  
  var strIndustry = document.getElementById('EIndustry')["value"].split("\n").join("<br />");
  var strELanguage = document.getElementById('ELanguage')["value"].split("\n").join("<br />");

  var strESDGCourse = document.getElementById('ESDGCourse')["value"].split("\n").join("<br />");


  var strESoftwarExp = document.getElementById('ESoftwarExp')["value"].split("\n").join("<br />");

  var strEMembership = document.getElementById('EMembership')["value"].split("\n").join("<br />");

  var strESKnowledge = document.getElementById('ESKnowledge')["value"].split("\n").join("<br />");

  var strShortbio = document.getElementById('Eshortbio')["value"].split("\n").join("<br />");

var insertObj={}
  try {
    if(IsAdminStaff)
    {
insertObj={
  Title: "SDG User Info",
  PersonalEmail: $("#personalmailID").val(),
  MobileNo: mobNumUpdate,
  HomeNo: homeNumUpdate,
  EmergencyNo: emergencyNumUpdate,
  // OfficeNo: officeNumUpdate,
  HomeAddLine: $("#PAddLineE").val(),
  HomeAddCity: $("#PAddCityE").val(),
  HomeAddState: $("#PAddStateE").val(),
  HomeAddPCode: $("#PAddPCodeE").val(),
  HomeAddCountry: $("#PAddCountryE").val(),
  IndustryExp: strIndustry,
  LanguageExp: strELanguage,
  SDGCourses: strESDGCourse,
  SoftwareExp: strESoftwarExp,
  Membership:strEMembership,
  SpecialKnowledge: strESKnowledge,
  Citizenship: $("#citizenshipE").val(),
  ShortBio: strShortbio,
  USDDailyRate: $("#USDDailyEdit").val(),
  // USDHourlyRate: $("#USDHourlyEdit").val(),
  EURDailyRate: $("#EURDailyEdit").val(),
  // EURHourlyRate: $("#EURHourlyEdit").val(),
  OtherCurrency: $("#othercurrDD").val(),
  ODailyRate: $("#ODailyEdit").val(),
  // OHourlyRate: $("#OHourlyEdit").val(),
  EffectiveDate: $("#EffectiveDateEdit").val()?$("#EffectiveDateEdit").val():null,
  BillingRateComments:$("#BillingRateCommentsEdit").val(), 
  signother: $("#significantOther").val(),
  children: $("#children").val(),
  WorkingSchedule: $("#workScheduleE").val(),
  SDGOfficeDetailsId : $("#workLocationDD").val(),
  StaffStatus: $("#staffstatusDD").val(),
  LinkedInLink:{
    "__metadata": { type: "SP.FieldUrlValue" },
    Description: "LinkedIn",
    Url: $("#linkedInID").val() 

  },
  Skype:$("#SkypeID").val(),
  stafffunction:$("#StaffFunctionEdit").val(),
  SDGAffiliation:$("#StaffAffiliatesEdit").val(),
  ///AssistantId: profileID,
  AssisstantName:$("#drpStaffAssitant option:selected").val(),
  ShowAvailability:$('#show-availability').prop('checked')
}
    }
    else
    {
      insertObj={
        Title: "SDG User Info",
        PersonalEmail: $("#personalmailID").val(),
        MobileNo: mobNumUpdate,
        HomeNo: homeNumUpdate,
        EmergencyNo: emergencyNumUpdate,
        // OfficeNo: officeNumUpdate,
        HomeAddLine: $("#PAddLineE").val(),
        HomeAddCity: $("#PAddCityE").val(),
        HomeAddState: $("#PAddStateE").val(),
        HomeAddPCode: $("#PAddPCodeE").val(),
        HomeAddCountry: $("#PAddCountryE").val(),
        IndustryExp: strIndustry,
        LanguageExp: strELanguage,
        SDGCourses: strESDGCourse,
        SoftwareExp: strESoftwarExp,
        Membership:strEMembership,
        SpecialKnowledge: strESKnowledge,
        Citizenship: $("#citizenshipE").val(),
        ShortBio: strShortbio,
        signother: $("#significantOther").val(),
        children: $("#children").val(),
        WorkingSchedule: $("#workScheduleE").val(),
        SDGOfficeDetailsId: $("#workLocationDD").val(),
        StaffStatus: $("#staffstatusDD").val(),
        LinkedInLink:{
          "__metadata": { type: "SP.FieldUrlValue" },
          Description: "LinkedIn",
          Url: $("#linkedInID").val()

        },
        Skype:$("#SkypeID").val(),
        stafffunction:$("#StaffFunctionEdit").val(),
        SDGAffiliation:$("#StaffAffiliatesEdit").val(),
        AssistantId: profileID,
      }
    }
    const update = await sp.web
    .getList(listUrl + "StaffDirectory")
    .items.getById(ItemID)
    .update(insertObj);
  alertify
  .alert("Submited Successfully", function(){
    alertify.message('OK');
    location.reload();
  });
  } catch (error) {
    ErrorCallBack(error, "EditItems");
  }
};
const editcancelFunction = () => {
  const viewDir = document.querySelector(".view-directory");
  const editDir = document.querySelector(".edit-directory");
  const editbtn = document.querySelector(".btn-edit");
  viewDir.classList.remove("hide");
  editDir.classList.add("hide");
  editbtn.classList.remove("hide");
//  $('#peoplepickerText').children().remove();
// sp-peoplepicker-editorInput
};
var inbetweendates = [];
const useravailabilityDetails = async() =>{
  var EndDateArr=[];
  var StartDateArr=[];
  var modifiedArray=[];
  var fDateArr=[];
  var NextArr=[];
  var StartArr=[];
  var EndArr=[];
  var availArr=[];
  var WeekArr=[];
  var FinalArr=[];

  var NextArr1=[];
  var StartArr1=[];
  var EndArr1=[];
  var availArr1=[];
  var WeekArr1=[];
  var FinalArr1=[];
  ///availList=AllAvailabilityDetails.filter((a)=>{return a.UserName.EMail.toLowerCase()==SelectedUserProfile[0].Usermail.toLowerCase()});

  availList=AllAvailabilityDetails.filter((a)=>{return a.UserEMail.toLowerCase()==SelectedUserProfile[0].Usermail.toLowerCase()});


  
  let availTableHtml = "";
  availList.forEach((avail)=>{
    availTableHtml  += `<tr><td>${avail.ProjectType}</td><td class="w100">${avail.Project?avail.Project:""}</td><td>${new Date(avail.StartDate).toLocaleDateString()}</td><td>${new Date(avail.EndDate).toLocaleDateString()}</td><td>${avail.Percentage}%</td><td class="w100">${avail.Comments?avail.Comments:""}</td><td><div class="d-flex"><div class="action-btn action-edit" data-toggle="modal" data-target="#addprojectmodal" data-id="${avail.ID}" id="editProjectAvailability"></div><div class="action-btn action-delete" data-id="${avail.ID}" id="deleteProjectAvailability"> </div></div></td></tr>`;
    var edate=new Date(new Date(avail.EndDate).toLocaleDateString()+" 00:00");
    var todayDate =new Date(new Date().toLocaleDateString()+" 00:00");
    if(edate>=todayDate&&avail.ProjectType!="Marketing"&&avail.ProjectType!="Overhead")
    {
      EndDateArr.push(new Date(avail.EndDate));
      StartDateArr.push(new Date(avail.StartDate));
      modifiedArray.push(avail);
    }
    
  });
  console.log(EndDateArr.sort(sortFunction));

  userAvailTable?userAvailTable.destroy():""
  $("#UserAvailabilityTbody").html("");
  $("#UserAvailabilityTbody").html(availTableHtml);

  // if(SelectedUserProfile[0].Availability!=0)
  // {
  //   $('.todayAvail').html(`<label>Available Now:(${SelectedUserProfile[0].Availability}% Available)</label>`)
  // }
  // else
  // {
  //   var EndDateArr =EndDateArr.sort(sortFunction);
  //   if(EndDateArr.length>0)
  //   {
  //     for(let i=0;i<EndDateArr.length; i++)
  //     {
  //       fDateArr =[];
  //       var endPercentage=0
  //       var leastEndDate =EndDateArr[i];
  //        modifiedArray.forEach((cal)=>{
  //         var edate=new Date(new Date(cal.EndDate).toLocaleDateString()+" 00:00");
  //         if(new Date(edate).toLocaleDateString()==new Date(leastEndDate).toLocaleDateString())
  //         fDateArr.push(cal)
  //       });
  //       fDateArr.map((b)=>b.Percentage?endPercentage+=parseInt(b.Percentage):endPercentage+=0);
  //       if(endPercentage<100){
  //         $('.todayAvail').html(`<label>Next available on :${new Date(leastEndDate).toLocaleDateString()}(${100-endPercentage}% Available)</label>`);
  //         break;
  //       }
  //       else if(i==EndDateArr.length-1)
  //       {
  //         var tomorrow = new Date(leastEndDate);
  //         tomorrow.setDate(tomorrow.getDate() + 1);
  //         $('.todayAvail').html(`<label>Next available on :${tomorrow.toLocaleDateString()}(${100}% Available)</label>`);
  //         break;
  //       }
  //     }

  //   }
  //   else{
  //     $('.todayAvail').html(`<label>Next available on :${new Date(leastEndDate).toLocaleDateString()}</label>`)
  //   }
  // }
  if(SelectedUserProfile[0].Availability!=0)
  {
    $('.NextAvail').html(" ")
    $('.todayAvail').html(`<label>Available Now:(${SelectedUserProfile[0].Availability}% Available)</label>`)
  }
  else
  {
    var EndDateArr =EndDateArr.sort(sortFunction);
    if(EndDateArr.length>0)
    {
        NextArr =[];
        StartArr=[];
        EndArr=[];
        availArr=[];
        WeekArr=[];
      for(let i=0;i<EndDateArr.length; i++)
      {
        fDateArr =[]; 
        
        var endPercentage=0
        var leastEndDate =EndDateArr[i];

         modifiedArray.forEach((cal)=>{
          var edate=new Date(new Date(cal.EndDate).toLocaleDateString()+" 00:00");
          var sdate=new Date(new Date(cal.StartDate).toLocaleDateString()+" 00:00");
        });

      }

      modifiedArray.forEach((cal)=>{
        var edate=new Date(new Date(cal.EndDate).toLocaleDateString()+" 00:00");
        var sdate=new Date(new Date(cal.StartDate).toLocaleDateString()+" 00:00");

        var Avainewday=new Date(new Date(cal.EndDate).toLocaleDateString()+" 00:00");
        Avainewday.setDate(Avainewday.getDate() + 1);
        NextArr.push(Avainewday);
        StartArr.push(sdate);
        EndArr.push(edate);

      });
      var checkFlag = false;
          for(var k=0;k<NextArr.length;k++){
            checkFlag = false;
            for(var j=0;j<StartArr.length;j++){
            if(StartArr[j]<=NextArr[k] && NextArr[k]&&NextArr[k]<=EndArr[j]){
              //console.log(availArr);
              checkFlag=true;
            }
            else{
                checkFlag=false;
                if(availArr.indexOf(NextArr[k])<0)
                availArr.push(NextArr[k]);
              
            }
            if(checkFlag)
            {
              if(availArr.indexOf(NextArr[k])>=0)
              {
                availArr.splice(availArr.indexOf(NextArr[k]),1)
                break;
              }
              else
              {
                break;
              }
            }          
            }
          }
                   console.log(availArr);
          availArr= availArr.sort(sortFunction);

          availArr = availArr.reduce(function (item, e1) {  
            var matches = item.filter(function (e2)  
            { return e1 == e2});  
            if (matches.length == 0) {  
                item.push(e1);  
            }  
            return item;  
        }, []);  
        console.log("availArr");
        console.log(availArr);

        for(var i=0;i<availArr.length;i++){

        if(new Date(availArr[i]).getDay()!=0 && new Date(availArr[i]).getDay()!=6)
        FinalArr.push(availArr[i]);
        else
        WeekArr.push(availArr[i]);

}
          if(FinalArr.length>0)
        {
          $('.todayAvail').html(`<label>Next 100% available on :${new Date(FinalArr[0]).toLocaleDateString()}</label>`)
          $('.NextAvail').html(" ")
        }
        else{
          if(new Date(WeekArr[0]).getDay()==0)
          {
            var Avainewday=new Date(new Date(WeekArr[0]).toLocaleDateString()+" 00:00");
            Avainewday.setDate(Avainewday.getDate() + 1);
            $('.todayAvail').html(`<label>Next 100% available on :${new Date(Avainewday).toLocaleDateString()}</label>`)
            $('.NextAvail').html(" ")
          }
          else if(new Date(WeekArr[0]).getDay()==6)
          {
            var Avainewday=new Date(new Date(WeekArr[0]).toLocaleDateString()+" 00:00");
            Avainewday.setDate(Avainewday.getDate() + 2);
            $('.todayAvail').html(`<label>Next 100% available on :${new Date(Avainewday).toLocaleDateString()}</label>`)
            $('.NextAvail').html(" ")
          }
        }
    }
    else{
      $('.todayAvail').html(`<label>Next available on :${new Date(leastEndDate).toLocaleDateString()}</label>`)
    }
  }
  //NextAvail
  if(SelectedUserProfile[0].Availability!=100 &&FinalArr.length==0 && WeekArr.length==0)
  {
    var EndDateArr1 =EndDateArr.sort(sortFunction);
    if(EndDateArr1.length>0)
    {
        NextArr1 =[];
        StartArr1=[];
        EndArr1=[];
        availArr1=[];
      for(let i=0;i<EndDateArr1.length; i++)
      {
        fDateArr =[]; 
        
        var endPercentage=0
        var leastEndDate =EndDateArr1[i];

         modifiedArray.forEach((cal)=>{
          var edate=new Date(new Date(cal.EndDate).toLocaleDateString()+" 00:00");
          var sdate=new Date(new Date(cal.StartDate).toLocaleDateString()+" 00:00");
        });

      }

      modifiedArray.forEach((cal)=>{
        var edate=new Date(new Date(cal.EndDate).toLocaleDateString()+" 00:00");
        var sdate=new Date(new Date(cal.StartDate).toLocaleDateString()+" 00:00");

        var Avainewday=new Date(new Date(cal.EndDate).toLocaleDateString()+" 00:00");
        Avainewday.setDate(Avainewday.getDate() + 1);
        NextArr1.push(Avainewday);
        StartArr1.push(sdate);
        EndArr1.push(edate);
      });
      var checkFlag = false;
          for(var k=0;k<NextArr1.length;k++){
            checkFlag = false;
            for(var j=0;j<StartArr1.length;j++){
            if(StartArr1[j]<=NextArr1[k] && NextArr1[k]&&NextArr1[k]<=EndArr1[j]){
              //console.log(availArr);
              checkFlag=true;
            }
            else{
                checkFlag=false;
                if(availArr1.indexOf(NextArr1[k])<0)
                availArr1.push(NextArr1[k]);
              
            }
            if(checkFlag)
            {
              if(availArr1.indexOf(NextArr1[k])>=0)
              {
                availArr1.splice(availArr1.indexOf(NextArr1[k]),1)
                break;
              }
              else
              {
                break;
              }
            }          
            }
          }
          console.log(availArr1);
          availArr1= availArr1.sort(sortFunction);
          availArr1 = availArr1.reduce(function (item, e1) {  
            var matches = item.filter(function (e2)  
            { return e1 == e2});  
            if (matches.length == 0) {  
                item.push(e1);  
            }  
            return item;  
        }, []);  
        console.log("availArr1");
        console.log(availArr1);

        for(var i=0;i<availArr1.length;i++){

        if(new Date(availArr1[i]).getDay()!=0 && new Date(availArr1[i]).getDay()!=6)
        FinalArr1.push(availArr1[i]);
        else
        WeekArr1.push(availArr1[i]);

}
          if(FinalArr1.length>0)
        {
          $('.NextAvail').html(`<label>Next 100% available on :${new Date(FinalArr1[0]).toLocaleDateString()}</label>`)
          
        }
        else{
          if(new Date(WeekArr1[0]).getDay()==0)
          {
            var Avainewday=new Date(new Date(WeekArr1[0]).toLocaleDateString()+" 00:00");
            Avainewday.setDate(Avainewday.getDate() + 1);
            $('.NextAvail').html(`<label>Next 100% available on :${new Date(Avainewday).toLocaleDateString()}</label>`)
            
          }
          else if(new Date(WeekArr1[0]).getDay()==6)
          {
            var Avainewday=new Date(new Date(WeekArr1[0]).toLocaleDateString()+" 00:00");
            Avainewday.setDate(Avainewday.getDate() + 2);
            $('.NextAvail').html(`<label>Next 100% available on :${new Date(Avainewday).toLocaleDateString()}</label>`)
            
          }
        }

    }
    else{
      $('.NextAvail').html(`<label>Next available on :${new Date(leastEndDate).toLocaleDateString()}</label>`)
    }
  }

      if(IsAdminStaff||(SelectedUserProfile[0].Usermail.toLowerCase()==currentMail.toLowerCase()&&IsgeneralStaff))
      {
        var Poptions = {
          destroy:true,
          "pageLength": 5,
          "order": [],
          language: {
            "emptyTable": "No data available"
          },
          columnDefs : [
            //hide the second & fourth column
            { 'visible': true, 'targets': [5] },
            { type: 'date', 'targets': [2] },
            { type: 'date', 'targets': [3] }
          ]
        };
      }
      else
    {
      var Poptions = {
        destroy:true,
        "pageLength": 5,
        "order": [],  
        language: {
          "emptyTable": "No data available"
        },
        
      columnDefs : [
        //hide the second & fourth column
        { 'visible': false, 'targets': [5] },
        { type: 'date', 'targets': [2] },
        { type: 'date', 'targets': [3] }
      ]

      };
     
    }




    userAvailTable =  (<any>$("#UserAvailabilityTable")).DataTable(Poptions);
    $('#UserAvailabilityTable').css("width","");


}

var sortFunction = function (a, b) {
  return Date.parse(a) - Date.parse(b);
};

function removeSelectedfile(filename) {
  for (var i = 0; i < bioAttachArr.length; i++) {
    if (bioAttachArr[i].name == filename) {
      ///filesQuantity[i].remove();
      bioAttachArr.splice(i, 1);
      break;
    }
  }
}
const removeAvailProject = async (ID) =>{
  await sp.web.getList(listUrl + "SDGAvailability").items.getById(ID).delete().then(async()=>{
    location.reload();
  });
}
const multipleMobNo = () => {
  $("#mobileNoSec").append(
    `<div class="d-flex mobNumbers"><input type="text" class="mobNo" id="mobileno1"/><span class="removeMobNo remove-icon"></span></div>`
  );
};
const multipleHomeNo = () => {
  $("#homeNoSec").append(
    `<div class="d-flex homeNumbers"><input type="text" class="mobNo" id="homeno1"/><span class="removeHomeNo remove-icon"></span></div>`
  );
};
const multipleEmergencyNo = () => {
  $("#emergencyNoSec").append(
    `<div class="d-flex emergencyNumbers"><input type="text" class="mobNo" id="emergencyno1"/><span class="removeemergencyNo remove-icon"></span></div>`
  );
};
const multipleOfficeNo = () => {
  $("#officeNoSec").append(
    `<div class="d-flex officeNumbers"><select class="officeNoCode">${CCodeHtml}</select><input type="text" class="mobNo" id="officeno1"/><span class="removeOfficeNo remove-icon"></span></div>`
  );
};

const fillEditSection = (ID) =>{
   editID=ID;
   var editedData = availList.filter((e)=>{return e.Id == parseInt(ID)});
   if(editedData[0].ProjectType == "Vacation"|| editedData[0].ProjectType == "Marketing"|| editedData[0].ProjectType == "Overhead"){
    $('#pName-lbl').text("Title");
    $('#percent-label').text("Percentage");
    $('#projectName').val($("#projecttypeDD").val());
    //$('#projectName').prop("disabled",true);
    $("#practiceAreaDD").val("Select");
    $("#OtherPracticeAreaDiv").addClass('hide')
  $('#add-basic-project,#add-basic-date,#add-percent-project,#Projectcomments-sec').show();
  $('#add-client-project,#add-practice-project,#projectAvailNotes-sec').hide();
  

  }else{
    $('#pName-lbl').text("Project Name");
    $('#percent-label').text("% Time on Project");
    $("#practiceAreaDD").val("Select");
    $('#projectName').val("");
    $("#OtherPracticeArea").val("");
    //$('#projectName').prop("disabled",false);
    $("#OtherPracticeAreaDiv").addClass('hide')
    $('#add-basic-project,#add-basic-date,#add-percent-project,#Projectcomments-sec').show();
    $('#add-client-project,#add-practice-project,#projectAvailNotes-sec').show();
  }
   
   if(editedData.length>0){
    var Sfinalmonth: any = "";
    var Sfinalday: any = "";
      var Sdd = new Date(editedData[0].StartDate).getDate();
      Sdd < 10 ? (Sfinalday = "0" + Sdd) : (Sfinalday = Sdd);
      var Smm = new Date(editedData[0].StartDate).getMonth() + 1;
      Smm < 10 ? (Sfinalmonth = "0" + Smm) : (Sfinalmonth = Smm);
      var Syyyy = new Date(editedData[0].StartDate).getFullYear();
      var Sdateformat = Syyyy + "-" + Sfinalmonth + "-" + Sfinalday;

      var Efinalmonth: any = "";
      var Efinalday: any = "";
      var Edd = new Date(editedData[0].EndDate).getDate();
      Edd < 10 ? (Efinalday = "0" + Edd) : (Efinalday = Edd);
      var Emm = new Date(editedData[0].EndDate).getMonth() + 1;
      Emm < 10 ? (Efinalmonth = "0" + Emm) : (Efinalmonth = Emm);
      var Eyyyy = new Date(editedData[0].EndDate).getFullYear();
      var Edateformat = Eyyyy + "-" + Efinalmonth + "-" + Efinalday;
      if(editedData[0].ProjectArea=="Others, please specify")
      {
        $('#OtherPracticeAreaDiv').removeClass('hide')
      }
      else
      {
        $('#OtherPracticeAreaDiv').addClass('hide')

      }
    $("#projectName").val(editedData[0].Project);
    $("#projectStartDate").val(Sdateformat);
    $("#projectEndDate").val(Edateformat);
    $("#projectPercent").val(editedData[0].Percentage);
    $("#practiceAreaDD").val(editedData[0].ProjectArea);
    $("#client").val(editedData[0].Client);
    $("#projecttypeDD").val(editedData[0].ProjectType);
    $("#projectCode").val(editedData[0].ProjectCode);
    $("#ProjectLocation").val(editedData[0].ProjectLocation);
    $("#OtherPracticeArea").val(editedData[0].OtherProjectArea);
    $("#projectAvailNotes").val(editedData[0].Notes);
    $("#Projectcomments").val(editedData[0].Comments)
   }


}
const availSubmitFunc = async() =>{
var bwArray=[];
var isAllSuccess=true
var enteredPercentage=parseInt(<any>$("#projectPercent").val())
var pName=$("#projectName").val()
    if($("#projecttypeDD").val()!="Marketing"&&$("#projecttypeDD").val()!="Overhead")
    {
      var sDate:any=$("#projectStartDate").val()
      var eDate:any=$("#projectEndDate").val();
      var startd:any = new Date(sDate);
      var endd = new Date(eDate);
      var newend = endd.setDate(endd.getDate()+1);
      endd = new Date(newend);
      while(startd < endd){
        bwArray.push(new Date(startd).toLocaleDateString()+" 00:00")
        var newDate = startd.setDate(startd.getDate() + 1);
        startd = new Date(newDate);
      }
          for(let i=0;i<bwArray.length;i++)
          {
            let datearr=new Date(bwArray[i]);
            let filteredData=[];
         availList.filter((data)=>{
           var sDate:any=new Date(new Date(data.StartDate).toLocaleDateString()+" 00:00");
           var eDate:any=new Date(new Date(data.EndDate).toLocaleDateString()+" 00:00");
            if(sDate<=datearr && eDate>=datearr && data.ProjectType!="Overhead"&& data.ProjectType!="Marketing")
            {
              filteredData.push(data);
            }
          });
        var dayValue = filteredData.reduce((n, {Percentage}) => n + parseInt(Percentage), 0);
        var correlationPercentage=100-parseInt(dayValue);
        if(enteredPercentage<=correlationPercentage){
          isAllSuccess=true
          
        }
        else{
           alertify.alert("Not able to add ur percentage in this date : "+datearr.toLocaleDateString());
          isAllSuccess=false
          break;
        }
      }
    }


 let ProjectPercent = 0;

  if(isAllSuccess)
  {
    const submitProject = await sp.web.getList(listUrl + "SDGAvailability").items.add({
        ///UserNameId:SelectedUserProfile[0].UserId,
        
        UserEMail:SelectedUserProfile[0].Usermail,
        UserFirstName:SelectedUserProfile[0].FirstName,
        UserLastName:SelectedUserProfile[0].LastName,
        UserJobTitle:SelectedUserProfile[0].JobTitle,
        UserTitle:SelectedUserProfile[0].Name,

        Project:pName,
        StartDate:$("#projectStartDate").val(),
        EndDate:$("#projectEndDate").val(),
        Percentage:enteredPercentage.toString(),
        ProjectArea:$("#practiceAreaDD").val(),
        ProjectType:$("#projecttypeDD").val(),
        Client:$("#client").val(),
        ProjectCode:$("#projectCode").val(),
        ProjectLocation:$("#ProjectLocation").val(),
        OtherProjectArea:$("#OtherPracticeArea").val(),
        Notes:$("#projectAvailNotes").val(),
        Comments:$("#Projectcomments").val()
      })

  $("#projectName").val("")
  $("#projectStartDate").val("")
  $("#projectEndDate").val("")
  $("#projectPercent").val("")
  $("#practiceAreaDD").val("")
  $("#projecttypeDD").val("")
  $("#client").val("")
  $("#projectCode").val("")
  $("#ProjectLocation").val("")
  $("#OtherPracticeArea").val("")
  $("#projectAvailNotes").val("")
  $("#Projectcomments").val("")
  alertify
  .alert("Submited Successfully", function(){
    alertify.message('OK');
    location.reload();
  }); 
  }
  else{
    alertify.alert(`Check the value entered in the "Percentage" field. It currently exceeds  100% across projects`);
  }


}
const availUpdateFunc = () =>{
  var bwArray=[];
  var isAllSuccess=true
  var enteredPercentage=parseInt(<any>$("#projectPercent").val())
  var pName=$("#projectName").val()

  if($("#projecttypeDD").val()!="Marketing"&&$("#projecttypeDD").val()!="Overhead")
    {
        var sDate:any=$("#projectStartDate").val()
        var eDate:any=$("#projectEndDate").val();
        var startd:any = new Date(sDate);
        var endd = new Date(eDate);
        var newend = endd.setDate(endd.getDate()+1);
        endd = new Date(newend);
        while(startd < endd){
          bwArray.push(new Date(startd).toLocaleDateString()+" 00:00")
          var newDate = startd.setDate(startd.getDate() + 1);
          startd = new Date(newDate);
        }

      for(let i=0;i<bwArray.length;i++)
      {
        let datearr=new Date(bwArray[i]);
        let filteredData=[];
     availList.filter((data)=>{
          var sDate:any=new Date(new Date(data.StartDate).toLocaleDateString()+" 00:00");
       var eDate:any=new Date(new Date(data.EndDate).toLocaleDateString()+" 00:00");
        if(sDate<=datearr && eDate>=datearr&&data.Id!=parseInt(editID)&& data.ProjectType!="Overhead"&& data.ProjectType!="Marketing")
        {
          filteredData.push(data);
        }
      });
    var dayValue = filteredData.reduce((n, {Percentage}) => n + parseInt(Percentage), 0);
    var correlationPercentage=100-parseInt(dayValue);
    if(enteredPercentage<=correlationPercentage){
      isAllSuccess=true
    }
    else{
      alertify.alert("Not able to add ur percentage in this date : "+datearr.toLocaleDateString());
      isAllSuccess=false
      break;
    }
  }
}

  let ProjectPercent = 0;
  if(isAllSuccess)
  {
    const updateProject =  sp.web
    .getList(listUrl + "SDGAvailability").items.getById(AvailEditID).update(
      {
        Project:pName,
        StartDate:$("#projectStartDate").val(),
        EndDate:$("#projectEndDate").val(),
        Percentage:enteredPercentage.toString(),
        ProjectArea:$("#practiceAreaDD").val(),
        ProjectType:$("#projecttypeDD").val(),
        Client:$("#client").val(),
        ProjectCode:$("#projectCode").val(),
        ProjectLocation:$("#ProjectLocation").val(),
        OtherProjectArea:$("#OtherPracticeArea").val(),
        Notes:$("#projectAvailNotes").val(),
        Comments:$("#Projectcomments").val()
      }
    );
    $("#projectName").val("")
    $("#projectStartDate").val("")
    $("#projectEndDate").val("")
    $("#projectPercent").val("")
    $("#practiceAreaDD").val("")
    $("#projecttypeDD").val("")
    $("#client").val("")
    $("#projectCode").val("")
    $("#ProjectLocation").val("")
    $("#OtherPracticeArea").val("")
    $("#projectAvailNotes").val("")
    $("#Projectcomments").val("")
    AvailEditID = 0;
    AvailEditFlag = false;
    alertify
  .alert("Updated Successfully", function(){
    alertify.message('OK');
    location.reload();
  });
  }  else{
    alertify.alert(`Check the value entered in the "Percentage" field. It currently exceeds  100% across projects`);
  }

}

const getGroups = async () =>{
  await sp.web.currentUser.get().then((user) => {
  sp.web.siteUsers.getById(user.Id).groups.get()
  .then((groupsData) => {
            groupsData.forEach(group => {
                if(group.Title == "General Employees")
                IsgeneralStaff=true
                else if(group.Title == "Special Access Employee")
                IssplStaff=true;
                else if(group.Title == "Staff Directory Admin")
                IsAdminStaff=true
              });
              IsAdminStaff||IssplStaff?$('.billingRateNav').show():$('.billingRateNav').hide()

  });
  });
}
function mandatoryforaddaction()
{
  var isAllvalueFilled=true;
  alertify.set('notifier','position', 'top-center');

  if(!$("#projecttypeDD").val()||$("#projecttypeDD").val()=="Select")
  {
    alertify.error("Please Select Project Type");
    isAllvalueFilled=false;
    
  }
  // else if(!$("#projectName").val())
  // {
  //   alertify.error("Please enter Project Name");
  //   isAllvalueFilled=false;
    
  // }
  else if(!$("#projectStartDate").val())
  {
    alertify.error("Please select Start Date");
    isAllvalueFilled=false;
    
  }

  else if(!$("#projectEndDate").val())
  {
    alertify.error("Please select End Date");
    isAllvalueFilled=false;
    
  }
  else if($("#projectEndDate").val()<$("#projectStartDate").val())
  {
    alertify.error("End Date Must be greater than Start Date");
    isAllvalueFilled=false;
    
  }
  else if(!$("#projectPercent").val())
  {
    alertify.error("Please enter % Time on Project");
    isAllvalueFilled=false;
    
  }
  else if($("#projectPercent").val()>100)
  {
    alertify.error("Please enter Less than or equal to 100%");
    isAllvalueFilled=false;
    
  }
  // else if(!$("#client").val())
  // {
  //   alertify.error("Please enter Client");
  //   isAllvalueFilled=false;
    
  // }
  // else if(!$("#projectCode").val())
  // {
  //   alertify.error("Please enter Project Code");
  //   isAllvalueFilled=false;
    
  // }
  // else if(!$("#practiceAreaDD").val())
  // {
  //   alertify.error("Please enter Practice Area");
  //   isAllvalueFilled=false;
    
  // }
  // else if(!$("#ProjectLocation").val())
  // {
  //   alertify.error("Please enter Project Location");
  //   isAllvalueFilled=false;
    
  // }
  else if(!$("#OtherPracticeArea").val()&&$("#practiceAreaDD").val()=="Others, please specify")
  {
    alertify.error("Please enter Project Location");
    isAllvalueFilled=false; 
  }
  // else if(!$("#projectAvailNotes").val())
  // {
  //   alertify.error("Please enter Availability Notes");
  //   isAllvalueFilled=false;
    
  // }
  // else if(!$("#Projectcomments").val())
  // {
  //   alertify.error("Please enter Comments");
  //   isAllvalueFilled=false;
    
  // }
  return isAllvalueFilled;
}
async function ErrorCallBack(error, methodname) {
  try {
    var errordata = {
      Error: error.message,
      MethodName: methodname,
    };
    await sp.web.lists
      .getByTitle("ErrorLog")
      .items.add(errordata)
      .then(function (data) {
        $('.loader').hide();
        AlertMessage("Something went wrong.please contact system admin");
      });
  } catch (e) {
    $('.loader').hide();
    AlertMessage("Something went wrong.please contact system admin");
  }
}
function AlertMessage(strMewssageEN) {
  alertify
    .alert()
    .setting({
      label: "OK",

      message: strMewssageEN,

      onok: function () {
        window.location.href = "#";
      },
    })
    .show()
    .setHeader("<em>Confirmation</em> ")
    .set("closable", false);
}