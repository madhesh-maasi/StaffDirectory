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
SPComponentLoader.loadScript(
  "https://ajax.aspnetcdn.com/ajax/4.0/1/MicrosoftAjax.js"
);

import * as $ from "jquery";
import { sp } from "@pnp/sp/presets/all";
import { graph } from "@pnp/graph";

import "@pnp/graph/users";
import "@pnp/graph/photos";
import "@pnp/sp/profiles";
SPComponentLoader.loadScript(
  "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.4.min.js"
  // "https://code.jquery.com/jquery-3.5.1.js"
);

import "../../ExternalRef/CSS/style.css";

SPComponentLoader.loadCss(
  "https://cdn.jsdelivr.net/npm/select2@4.1.0-beta.1/dist/css/select2.min.css"
);
SPComponentLoader.loadCss(
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"
);
//import "datatables";

require("datatables.net-dt");
require("datatables.net-rowgroup-dt");
SPComponentLoader.loadCss(
  "https://cdn.datatables.net/1.10.24/css/jquery.dataTables.css"
);

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
export default class StaffdirectoryWebPart extends BaseClientSideWebPart<IStaffdirectoryWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then((_) => {
      sp.setup({
        spfxContext: this.context,
      });
      graph.setup({
        spfxContext: this.context,
      });
    });
  }

  public render(): void {
    listUrl = this.context.pageContext.web.absoluteUrl;
    var siteindex = listUrl.toLocaleLowerCase().indexOf("sites");
    listUrl = listUrl.substr(siteindex - 1) + "/Lists/";
    this.domElement.innerHTML = `
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
         <li><a href="#" class="OutConsStaffgrouping">By Staff Function</a></li>
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
         <li><a href="#" class="AffFirstname">By First Name</a></li>
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
       <div id="collapseSix" class="clsCollapse collapse" aria-labelledby="headingSix" data-parent="#accordionExample">
         <div class="card-body">
         <div class="filter-section">
         <ul>
         <li><a href="#" class="SDGOfficeInfoLastName">By Last Name</a></li>
         <li><a href="#"class="SDGOfficeInfoFirstName">By First Name</a></li>
         </ul>
         </div>
         </div>
       </div>
     </div>
     <div class="card">
       <div class="card-header nav-items StaffAvailability" id="headingSeven">
           <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseSeven" aria-expanded="false" aria-controls="collapseSeven"><span class="nav-icon staff-avail"></span>Staff Availability</div>
       </div>
       <div id="collapseSeven" class="clsCollapse collapse" aria-labelledby="headingSeven" data-parent="#accordionExample">
         <div class="card-body">
         <div class="filter-section">
         <!--<ul>
         <li><a href="#">By Office</a></li>
         <li><a href="#">By Title</a></li>
         </ul>-->
         </div>
         </div>
       </div>
     </div> 
     <div class="card">  
       <div class="card-header nav-items SDGBillingRate" id="headingEight">
           <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseEight" aria-expanded="false" aria-controls="collapseEight"><span class="nav-icon billing-rate"></span>Billing Rates</div>
       </div>
       <div id="collapseEight" class="clsCollapse collapse" aria-labelledby="headingEight" data-parent="#accordionExample">
         <div class="card-body">
         <div class="filter-section">
         <ul>
         <li><a href="#" class="SDGBillingRateLastName">By Last Name</a></li>
         <li><a href="#" class="SDGBillingRateFirstName">By First Name</a></li>
         </ul>
         </div>
         </div>
       </div>
     </div>
   </div> 
     </div>
     </div>
     <div class="right">
     <div class="sdh-employee" id="SdhEmployeeDetails">
     <!-- <div class="title-section">
     <h2>Overview</h2>
     </div> --> 
     <div class="title-filter-section">
     </div>   
     <div class="sdh-emp-table oDataTable">
     <table  id="SdhEmpTable">
     <thead>
     <tr>
     <th>Name</th>
     <th>First Name</th> 
     <th>Last Name</th>
     <th>Phone Number</th>
     <th>Location</th>
     <th>Job Title</th>
     <th>Title</th>
     <th>Assistant</th>
     </tr>
     </thead>
     <tbody id="SdhEmpTbody">
     </tbody>
     </table>
     </div> 
     <div class="sdh-outside-table oDataTable hide">
     <table  id="SdhOutsideTable">
     <thead>
     <tr>
     <th>Name</th>
     <th>First Name</th> 
     <th>Last Name</th>
     <th>Phone Number</th>
     <th>Location</th>
     <th>Job Title</th>
     <th>Title</th>
     <th>Assistant</th>
     </tr>
     </thead>
     <tbody id="SdhOutsideTbody">
     </tbody>
     </table>
     </div> 
     <div class="sdh-Affilate-table oDataTable hide">
     <table  id="SdhAffilateTable">
     <thead>
     <tr>
     <th>Name</th>
     <th>First Name</th> 
     <th>Last Name</th>
     <th>Phone Number</th>
     <th>Location</th>
     <th>Job Title</th>
     <th>Title</th>
     <th>Assistant</th>
     </tr>
     </thead>
     <tbody id="SdhAffilateTbody">
     </tbody>
     </table>
     </div>

     <div class="sdh-Allumni-table oDataTable hide">
     <table  id="SdhAllumniTable">
     <thead>
     <tr>
     <th>Name</th>
     <th>First Name</th> 
     <th>Last Name</th>
     <th>Phone Number</th>
     <th>Location</th>
     <th>Job Title</th>
     <th>Title</th>
     <th>Assistant</th>
     </tr>
     </thead>
     <tbody id="SdhAllumniTbody">
     </tbody>
     </table>
     </div>
     <div class="sdh-AllPeople-table oDataTable hide">
     <table  id="SdhAllPeopleTable">
     <thead>
     <tr>
     <th>Name</th>
     <th>First Name</th> 
     <th>Last Name</th>
     <th>Phone Number</th>
     <th>Location</th>
     <th>Job Title</th>
     <th>Title</th>
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
     <th>Address</th>
     </tr>
     </thead>
     <tbody id="SdgofficeinfoTbody">
     </tbody>
     </table>
     </div>
     <div class="sdgbillingrateTable oDataTable hide">
     <table  id="SdgBillingrateTable">
     <thead>
     <tr>
     <th>Name</th>
     <th>Satff Function</th> 
     <th>Daily Rate</th>
     <th>Hourly Rate</th>
     <th>Effective Date</th>
     </tr>
     </thead>
     <tbody id="SdgBillingrateTbody">
     </tbody>
     </table>
     </div>
     <div class="StaffAvailabilityTable oDataTable hide">
     <table id="StaffAvailabilityTable">
     <thead>
     <tr><th>User</th><th>Location</th><th>Staff Affiliates</th><th>Availability</th></tr>
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
     <img src="data:image/png;base64,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" class="profile-picture"> 
     </div>  
     <div class="profile-name-section">
     <p class="profile-user-name" id="UserProfileName">Sample User</p>
     <p class="profile-user-mail" id="UserProfileEmail"><span class="user-mail-icon"></span>Sample mail</p>        
     </div>   
     </div>
     <div class="user-details-section">
     <div class="profile-details-left">
     <div class="user-info">
     <label>Job Title:</label>
     <div class="title-font" id="user-job-title"></div>
     </div>
     <div class="user-info">   
     <label>SDG Affiliation :</label> 
     <div class="title-font" id="user-Designation"></div>
     </div>
     <div class="user-info">   
     <label>Staff Function :</label> 
     <div class="title-font" id="user-staff-function"></div>
     </div>
     
     </div>
     <div class="profile-details-right">
     <div class="user-info"> 
     <label>Mobile:</label>
     <div class="title-font" id="user-phone"></div>
     </div>
     
     <div class="user-info hide"><label>Personal Mail :</label><div class="title-font" id="userpersonalmail"></div></div> 
     <div class="user-info">
      <div class="d-flex align-item-center"><label>LinkedIn :</label><div class="" id="linkedinIDview"></div></div>
      </div>
     </div> 
     </div>
     </div>
     <div class="user-profile-tabs">
     <div class="tab-section"> 
     <div class="tab-header-section">  
     <ul class="nav nav-tabs">
       <li class="active"><a data-toggle="tab" href="#home">Directory Information</a></li>
       <li id="availabilityTab"><a data-toggle="tab" href="#menu1">Availability</a></li>
     </ul>
     
     </div>
     <div>
     <div class="tab-content">
    <div id="home" class="tab-pane fade in active">
    <div class="text-right" ><button class="btn btn-edit" id="btnEdit">Edit</button></div> 
      <div id="DirectoryInformation" class="d-flex view-directory">
      <div class="DInfo-left col-6">
      <div class="work-address">
      <h4>Work Address</h4>
      <div class="d-flex"><label>Location :</label><div class="address-details lblRight" id="WLoctionDetails"></div></div>
      <div class="d-flex align-item-center"><label>Address:</label><div class="address-details lblRight" id="WAddressDetails"></div>
      </div> 
      </div>
      
      <div class="Assistant-view" id="viewAssistant">
      
      </div>
      <div class="personal-info">
      <h4>Personal Info</h4>
      <div class="address-details" id="PersonaInfo"> 
      <div class="d-flex"><label>Address Line:</label><div id="PAddLine" class="lblRight"></div></div>
      <div class="d-flex"><label>City:</label><div id="PAddCity" class="lblRight"></div></div>
      <div class="d-flex"><label>State:</label><div id="PAddState" class="lblRight"></div></div>
      <div class="d-flex"><label>Postal Code :</label><div id="PAddPCode" class="lblRight"></div></div> 
      <div class="d-flex"><label>Country:</label><div id="PAddPCountry" class="lblRight"></div></div>
      <div class="d-flex"><label>Significant Other :</label><div id="PSignOther" class="lblRight"></div></div>
      <div class="d-flex"><label>Children :</label><div id="PChildren" class="lblRight"></div></div>
      </div>
      </div>
      
      
      
      <div class="StaffStatus">
      <h4>Staff Status</h4>
      <p class="lblRight" id="staffStatus"></p> 
      <div id="workscheduleViewSec">
      <div class="d-flex"><label>Work Schedule</label><p class="lblRight" id="workSchedule"></p></div>
      
      </div>
      </div>
      <div class="citizen-info">
      
      <div class="address-details" id="CitizenInfo"> 
      <div><label>Nationality :</label><label id="citizenship" class="lblRight"></label></div>
      </div>
      </div>
      </div>
      <div class="DInfo-right col-6">
      <div class="user-billing-rates">
      <h4>Billing Rates</h4>
      <div id="BillingRateDetails">
      <div class="billing-rates"><label>USD Daily Rates</label><div class="usd-daily-rate" id="UsdDailyRate"></div></div>
      <div class="billing-rates"><label>USD Hourly Rates</label><div class="usd-hourly-rate" id="UsdHourlyRate"></div></div>
      <div class="billing-rates"><label>EUR Daily Rates</label><div class="eur-daily-rate" id="EURDailyRate"></div></div>
      <div class="billing-rates"><label>EUR Hourly Rates</label><div class="eur-hourly-rate" id="EURHourlyRate"></div></div>
      <div class="billing-effective-date"><label>Effective Date</label><div class="effective-date" id="EffectiveDate"></div></div>
      </div>
      </div>
      <div class="Biography-Experience"> 
      <h4>Biography and Experience</h4>
      <div class="address-details" id="BioExp">  
      <h5>Short Bio</h5>
      <p id="shortbio" class="lblRight"></p> 
      <h5>Bio Attachment(s)</h5>
      <div class="bio-attachment-section" id="bioAttachment"></div>
      <div class="other-exp">
      <h5>Other Experience Details</h5> 
      <div class="exp">
      <div class=""><label>Industries</label>
      <p id="IndustryExp" class="lblRight"></p>
      </div>
      <div class=""><label>Languages</label>
      <p id="LanguageExp" class="lblRight"></p>
      </div>
      </div>
      <div class="exp">
      <div class=""><label>SDG Courses</label>
      <p id="SDGCourse" class="lblRight"></p>
      </div>
      <div class=""><label>Software</label>
      <p id="SoftwareExp" class="lblRight"></p>
      </div>
      </div>
      <div class="exp">
      <div class=""><label>Memberships</label>
      <p id="MembershipExp" class="lblRight"></p>
      </div>
      <div class=""><label>Special Knowledge</label>
      <p id="SpecialKnowledge" class="lblRight"></p>
      </div>
      </div>
      </div>
      </div>
      </div>
      </div>  
      </div> 
      <div id="DirectoryInformationEdit" class="edit-directory hide">
      <div class="d-flex">
      <div class="DInfo-left col-6">
      <div class="work-address">
      <h4>Work Address</h4>
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
      <label>Staff Function</label>
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
      <div class="w-100"><div id="peoplepickerText" title="APickerField"></div></div>
      
      </div>
      </div>
      <div class="contact-info">
      <h4>Contact Info</h4>
      <div class="address-details" id="ContactInfo"> 
      <div class="d-flex"><label>Personal Mail :</label><div class="w-100"><input type="text" id="personalmailID"></div></div>
      <div class="d-flex"><label>Mobile No :</label><div class="w-100" id ="mobileNoSec"><div class="d-flex mobNumbers"><select class="mobNoCode"></select><input type="number" class="mobNo" id="mobileno1"/><span class="addMobNo add-icon"></span></div></div></div>
      <div class="d-flex"><label>Home No :</label><div class="w-100" id="homeNoSec"><div class="d-flex homeNumbers"><select class="homeNoCode"></select><input type="number" class="homeno" id="homeno"/><span class="addHomeNo add-icon"></span></div></div></div>
      <div class="d-flex"><label>Emergency No :</label><div class="w-100" id="emergencyNoSec"><div class="d-flex emergencyNumbers"><select class="emergencyNoCode"></select><input type="number" class="emergencyno" id="emergencyno" /><span class="addEmergencyNo add-icon"></span></div></div></div>
      
      <div class="d-flex"><label>Significant Other :</label><div class="w-100"><textarea id="significantOther"></textarea></div></div>
      <div class="d-flex"><label>Children :</label><div class="w-100"><textarea id="children"></textarea></div></div>
      <div class="d-flex"><label>LinkedIn ID :</label><div class="w-100"><input type="text" id="linkedInID"></div></div>
      </div> 
      </div>
      <div class="personal-info">
      <h4>Personal Info</h4> 
      <div class="address-details" id="PersonaInfo"> 
      <div class="d-flex"><label>Address Line:</label><div class="w-100"><input type="text" id="PAddLineE"></div></div>
      <div class="d-flex"><label>City:</label><div class="w-100"><input type="text" id="PAddCityE"></div></div>
      <div class="d-flex"><label>State:</label><div class="w-100"><input type="text" id="PAddStateE"></div></div>
      <div class="d-flex"><label>Postal Code :</label><div class="w-100"><input type="text" id="PAddPCodeE"></div></div> 
      <div class="d-flex"><label>Country:</label><div class="w-100"><input type="text" id="PAddCountryE"></div></div>
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
      <div class="d-flex w-100"><label>Nationality:</label><div class="w-100"><input type="text" id="citizenshipE"></div></div>
      </div>
      </div>
      </div> 

      <div class="DInfo-right col-6">
      <div class="user-billing-rates">
      <h4>Billing Rates</h4>

      <div id="BillingRateDetailsEdit">
      <div class="billing-rates"><label>USD Daily Rates</label><div class="usd-daily-rate"></div><input type="number" id="USDDailyEdit"/></div>
      <div class="billing-rates"><label>USD Hourly Rates</label><div class="usd-hourly-rate"></div><input type="number" id="USDHourlyEdit" disabled/></div>
      <div class="billing-rates"><label>EUR Daily Rates</label><div class="eur-daily-rate"></div><input type="number" id="EURDailyEdit"/></div>
      <div class="billing-rates"><label>EUR Hourly Rates</label><div class="eur-hourly-rate"></div><input type="number" id="EURHourlyEdit" disabled/></div>
      <div class="billing-rates"><label>Other Currency</label><div class="eur-hourly-rate"></div><select id="othercurrDD"></select></div>
      <div class="billing-rates"><label>Daily Rate</label><div class="eur-hourly-rate"></div><input type="number" id="ODailyEdit"/></div>
      <div class="billing-rates"><label>Hourly Rate</label><div class="eur-hourly-rate"></div><input type="number" id="OHourlyEdit" disabled/></div>
      <div class="billing-effective-date"><label>Effective Date</label><div class="effective-date"><input type="date" id="EffectiveDateEdit"/></div></div>
      </div>
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
      <div class=""><label>SDG Courses</label>
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
      </div>
      <div class="availability-table-section">
      <table id="UserAvailabilityTable">
      <thead>
      <tr>
      <th>Project Name</th>
      <th>Start Date</th>
      <th>End Date</th>
      <th>Percentage</th>
      <th>Comments</th>
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
      <div class="d-flex">
      <div class="d-flex col-6"><label>Project Type</label><div class="w-100"><select id="projecttypeDD"><option value="sample">Sample</option></select></div></div>
      <div class="d-flex col-6"><label>Project Name</label><div class="w-100"><input type="text" id="projectName" /></div></div>
      </div>
      <div class="d-flex">
      <div class="d-flex col-6"><label>Start Date</label><div class="w-100"><input type="date" id="projectStartDate" /></div></div>
        <div class="d-flex col-6"><label>End Date</label><div class="w-100"><input type="date" id="projectEndDate" /></div></div>
      </div>   
        <div class="d-flex">
        <div class="d-flex col-6"><div id="percentageDiv" class="d-flex w-100"><label>Percent on Project</label><div class="w-100"><input type="text" id="projectPercent" /></div></div></div>
        
        </div>
        
        <div class="d-flex">
        <div class="d-flex col-6"><label>Client</label><div class="w-100"><input type="text" id="client" /></div></div>
        <div class="d-flex col-6"><label>Project Code</label><div class="w-100"><input type="text" id="projectCode" /></div></div>
        </div>
        
        <div class="d-flex">
        <div class="d-flex col-6"><label>Practice Area</label><div class="w-100"><select id="practiceAreaDD"><option value="sample">Sample</option></select></div></div>
        <div class="d-flex col-6"><label>Project Location</label><div class="w-100"><input type="text" id="ProjectLocation" /></div></div>
        </div>
        <div class="d-flex">
        <div class="d-flex col-6"><label>Availability Notes</label><div class="w-100"><textarea id="projectAvailNotes" ></textarea></div></div>
        <div class="d-flex col-6"><label>Comments</label><div class="w-100"><textarea id="Projectcomments"></textarea></div></div></div>
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

    const username = document.querySelectorAll(".usernametag");
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
    ActiveSwitch();
    $(".SDHEmployee").click(() => {
      SelectedUserProfile = [];
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
          lengthMenu: [50, 100],
        };
        
        bindEmpTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          lengthMenu: [50, 100],
        };
        bindEmpTable(options);
      }
    });
    $(".OutsidConsultant").click(() => {
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
          lengthMenu: [50, 100],
        };
        bindOutTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          lengthMenu: [50, 100],
        };
        bindOutTable(options);
      }
    });
    $(".SDHAffiliates").click(() => {
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
          lengthMenu: [50, 100],
        };
        bindAffTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          lengthMenu: [50, 100],
        };
        bindAffTable(options);
      }
    });
    $(".SDHAlumini").click(() => {
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
          lengthMenu: [50, 100],
        };
        bindAlumTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
          lengthMenu: [50, 100],
        };
        bindAlumTable(options);
      }
    });
    $(".SDHShowAll").click(() => {
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
        };
        bindAllDetailTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindAllDetailTable(options);
      }
    });
    $(".SDGOfficeInfo").click(() => {
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
        };
        bindOfficeTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindOfficeTable(options);
      }
    });

  }
    // Employee Filters
    $(".sdhLocgrouping").click(() => {
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
      };
      bindEmpTable(options);
    });
    $(".sdhlastnamesort").click(() => {
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
      };
      bindEmpTable(options);
    });
    //OutSideConsultant
    $(".OutConslastnamesort").click(() => {
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
      };
      bindOutTable(options);
    });
    $(".OutConsFirstnamesort").click(() => {
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
      };
      bindOutTable(options);
    });
    $(".OutConsLocgrouping").click(() => {
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
      };
      bindAffTable(options);
    });
    $(".AffFirstnamesort").click(() => {
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
      };
      bindStaffAvailTable(options);
    });
    // Allumni
    $(".SDHAlumniLastName").click(() => {
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
      };
      bindAlumTable(options);
    });
    $(".SDHAlumniFirstName").click(() => {
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
      };
      bindAlumTable(options);
    });
    $(".SDHAlumniOffice").click(() => {
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
      };
      bindAllDetailTable(options);
    });
    $(".SDHShowAllFirstName").click(() => {
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
      };
      bindAllDetailTable(options);
    });
    $(".StaffAvailability").click(() => {
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
      };
      bindAllDetailTable(options);
    });
    $(".SDGBillingRate").click(() => {
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
      };
      bindBillingRateTable(options);
    });
    $(".SDGBillingRateLastName").click(() => {
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
      };
      bindBillingRateTable(options);
    });
    $(".SDGBillingRateFirstName").click(() => {
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
      };
      bindBillingRateTable(options);
    });
    $(".SDGOfficeInfoFirstName").click(() => {
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
        };
        bindOfficeTable(options);
    });$(".SDGOfficeInfoLastName").click(() => {
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
        };
        bindOfficeTable(options);
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
      if(AvailEditFlag){
        availUpdateFunc();
      }else{
        availSubmitFunc();
      }
      
    })
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
      // console.log(bioAttachArr);
    });
    $(document).on("click", ".clsRemove", function () {
      // console.log(bioAttachArr);
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
      removeAvailProject(parseInt(AItemID));
      e.currentTarget.parentElement.parentElement.parentElement.remove()
    })

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
        (add) => $("#workLocationDD").val() == add.OfficePlace
      )[0].OfficeFullAdd);
    })
    $(document).on("change","#projecttypeDD",()=>{
      if($("#projecttypeDD").val() == "Billable/Client" || $("#projecttypeDD").val() == "Marketing"){
       
          $("#percentageDiv").removeClass("hide")
        
      }else{
        $("#projectPercent").val("");
          $("#percentageDiv").addClass("hide")
       
      }
    })
    $(document).on("click","#editProjectAvailability",(e)=>{
      AvailEditFlag = true;
      let AEditItemID = e.currentTarget.getAttribute("data-id");
      AvailEditID = AEditItemID;
      fillEditSection(AvailEditID);
    });
    $(document).on("click","#closeModal",()=>{
      AvailEditFlag = false;
      AvailEditID = 0;
      $("#projectName").val("")
      $("#projectStartDate").val("")
      $("#projectEndDate").val("")
      $("#projectPercent").val("")
      $("#practiceAreaDD").val("")
      $("#client").val("")
      $("#projectCode").val("")
      $("#ProjectLocation").val("")
      $("#projectAvailNotes").val("")
      $("#Projectcomments").val("")
    })
    
    $(document).on("click",".usernametag",()=>{
      if(SelectedUserProfile[0].Affiliation != "Employee" && SelectedUserProfile[0].Affiliation != "Outside Consultant"){
        $("#menu1").addClass("hide");
        $("#availabilityTab").addClass("hide");
      }else{
        $("#menu1").removeClass("hide");
        $("#availabilityTab").removeClass("hide");
      }
    })
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
  let StaffStatusDD = document.querySelector("#staffstatusDD");
  let LocationDD = document.querySelector("#workLocationDD");
  let othercurrDD = document.querySelector("#othercurrDD");
  let StaffFunctionDD = document.querySelector("#StaffFunctionEdit");
  let StaffAffiliatesDD = document.querySelector("#StaffAffiliatesEdit");
  let AvailProjectTypeDD = document.querySelector("#projecttypeDD");
  let AvailPracticeAreaDD = document.querySelector("#practiceAreaDD");
  let LocOptionHtml = "";
  let staffOptionHtml = "";
  let otherCurrHtml = "";
  let StaffFunHtml = "";
  let StaffAffHtml = "";
  let AvailProjTypeHtml = "";
  let AvailPracAreaDD = "";
  // let CCodeHtml = "";
  let listLocation = await sp.web
    .getList(listUrl + "StaffDirectory")
    .fields.filter("EntityPropertyName eq 'SDGOffice'")
    .get();
  let listStaffStatus = await sp.web
    .getList(listUrl + "StaffDirectory")
    .fields.filter("EntityPropertyName eq 'StaffStatus'")
    .get();
  let listOtherCurr = await sp.web
    .getList(listUrl + "StaffDirectory")
    .fields.filter("EntityPropertyName eq 'OtherCurrency'")
    .get();
  let CountryCode = await sp.web
    .getList(listUrl + "StaffDirectory")
    .fields.filter("EntityPropertyName eq 'CountryCode'")
    .get();
  let listStaffFunction = await sp.web
  .getList(listUrl + "StaffDirectory")
  .fields.filter("EntityPropertyName eq 'stafffunction'")
  .get();
  let listStaffAff = await sp.web
  .getList(listUrl + "StaffDirectory")
  .fields.filter("EntityPropertyName eq 'SDGAffiliation'")
  .get();
  let AvailProjectType = await sp.web
  .getList(listUrl + "SDGAvailability")
  .fields.filter("EntityPropertyName eq 'ProjectType'")
  .get();
  let AvailPracticeArea = await sp.web
  .getList(listUrl + "SDGAvailability")
  .fields.filter("EntityPropertyName eq 'ProjectArea'")
  .get();

  AvailProjectType[0]["Choices"].forEach((type) => {
    AvailProjTypeHtml += `<option value="${type}">${type}</option>`;
  });
    AvailPracticeArea[0]["Choices"].forEach((Area) => {
      AvailPracAreaDD += `<option value="${Area}">${Area}</option>`;
  });


  listLocation[0]["Choices"].forEach((li) => {
    LocOptionHtml += `<option value="${li}">${li}</option>`;
  });
  listStaffStatus[0]["Choices"].forEach((stff) => {
    staffOptionHtml += `<option value="${stff}">${stff}</option>`;
  });
  listOtherCurr[0]["Choices"].forEach((curr) => {
    otherCurrHtml += `<option value="${curr}">${curr}</option>`;
  });
  CountryCode[0]["Choices"].forEach((CCode) => {
    CCodeArr.push(CCode);
    CCodeHtml += `<option value="${CCode}">${CCode}</option>`;
  });
  listStaffFunction[0]["Choices"].forEach((func) => {
    StaffFunHtml += `<option value="${func}">${func}</option>`;
  });
  listStaffAff[0]["Choices"].forEach((Aff) => {
    StaffAffHtml += `<option value="${Aff}">${Aff}</option>`;
  });
  AvailProjectTypeDD.innerHTML = AvailProjTypeHtml;
  AvailPracticeAreaDD.innerHTML = AvailPracAreaDD;
  LocationDD.innerHTML = LocOptionHtml;
  StaffStatusDD.innerHTML = staffOptionHtml;
  othercurrDD.innerHTML = otherCurrHtml;
  StaffFunctionDD.innerHTML = StaffFunHtml;
  StaffAffiliatesDD.innerHTML = StaffAffHtml;

  $(".mobNoCode,.homeNoCode,.emergencyNoCode").html(CCodeHtml);
  await sp.web
    .getList(listUrl + "StaffDirectory")
    .items.select(
      "*",
      "User/EMail",
      "User/Title",
      "User/FirstName",
      "User/LastName",
      "User/JobTitle",
      "Assistant/EMail",
      "Assistant/Title"
    )
    .expand("User,Assistant")
    .get()
    .then((listitem: any) => {
      console.log(listitem);
      listitem.forEach((li) => {
        UserDetails.push({
          Name: li.User.Title != null ? li.User.Title : "Not Available",
          FirstName: li.User.FirstName != null ? li.User.FirstName : "Not Available",
          LastName: li.User.LastName != null ? li.User.LastName  : "Not Available",
          Usermail: li.User.EMail != null ? li.User.EMail : "Not Available",
          UserPersonalMail: li.PersonalEmail != null ? li.PersonalEmail : "Not Available",
          JobTitle: li.User.JobTitle != null ? li.User.JobTitle : "Not Available",
          Assistant: li.Assistant.Title != null ? li.Assistant.Title : "Not Available",
          AssistantMail: li.Assistant.EMail != null ? li.Assistant.EMail : "Not Available",
          PhoneNumber: li.MobileNo != null ? li.MobileNo : "Not Available",
          Location: li.SDGOffice != null ? li.SDGOffice : "Not Available",
          Title: li.stafffunction != null ? li.stafffunction : "Not Available",
          Affiliation: li.SDGAffiliation != null ? li.SDGAffiliation : "Not Available",
          HAddLine: li.HomeAddLine != null ? li.HomeAddLine : "Not Available",
          HAddCity: li.HomeAddCity != null ? li.HomeAddCity : "Not Available",
          HAddState: li.HomeAddState != null ? li.HomeAddState : "Not Available",
          HAddPCode: li.HomeAddPCode != null ? li.HomeAddPCode : "Not Available",
          HAddPCountry: li.HomeAddCountry != null ? li.HomeAddCountry : "Not Available",
          ShortBio: li.ShortBio != null ? li.ShortBio : "Not Available",
          Citizen: li.Citizenship != null ? li.Citizenship : "Not Available",
          Industry: li.IndustryExp != null ? li.IndustryExp : "Not Available",
          Language: li.LanguageExp != null ?li.LanguageExp : "Not Available",
          SDGCourse: li.SDGCourses != null ? li.SDGCourses : "Not Available",
          Software: li.SoftwareExp != null ? li.SoftwareExp : "Not Available",
          Membership: li.Membership != null ? li.Membership : "Not Available",
          SpecialKnowledge: li.SpecialKnowledge != null ? li.SpecialKnowledge : "Not Available",
          USDDaily: li.USDDailyRate,
          USDHourly: li.USDHourlyRate,
          EURDaily: li.EURDailyRate,
          EURHourly: li.EURHourlyRate,
          OtherCurr: li.OtherCurrency,
          OtherCurrDaily: li.ODailyRate,
          OtherCurrHourly: li.OHourlyRate,
          EffectiveDate: li.EffectiveDate != null ? li.EffectiveDate : "Not Available",
          StaffStatus: li.StaffStatus != null ? li.StaffStatus : "Not Available",
          WorkSchedule: li.WorkingSchedule != null ? li.WorkingSchedule : "Not Available",
          ItemID: li.ID != null ? li.ID : "Not Available",
          LinkedInID: li.LinkedInLink != null ? li.LinkedInLink : "Not Available",
          SignOther: li.signother != null ? li.signother : "Not Available",
          Child: li.children != null ? li.children : "Not Available",
          HomeNo: li.HomeNo != null ? li.HomeNo : "Not Available",
          EmergencyNo: li.EmergencyNo != null ? li.EmergencyNo : "Not Available",
          
        });
      });
      getTableData();
      console.log(UserDetails);
    });
  // UserProfileDetail();
};
const ActiveSwitch = () => {
  let navItems = document.querySelectorAll(".nav-items");
  navItems.forEach((li) => {
    li.addEventListener("click", (e) => {
      let activeClass = document.querySelectorAll(".nav-items");
      activeClass.forEach((activeC) => {
        activeC["classList"].remove("show");
      });
      // console.log(e.currentTarget);
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
  let OfficeDetails = await sp.web
    .getList(listUrl + "SDGOfficeInfo") 
    .items.get();
  // console.log(OfficeDetails);
  OfficeDetails.forEach((oDetail) => {
    OfficeTable += `<tr><td>${oDetail.Office}</td><td>${
      oDetail.Phone != "null" ? oDetail.Phone.split("^").join("</br>") : ""
    }</td><td>${
      oDetail.Address != "null" ? oDetail.Address.split("^").join("</br>") : ""
    }</td></tr>`;
  });
  let AvailHtml = "";
let AllAvailabilityDetails = await sp.web
.getList(listUrl + "SDGAvailability")
.items
.get();

let AvalabilityUsers = AllAvailabilityDetails.map((users)=>{
  return (
    users.UserName
  );
});

AvalabilityUsers = AvalabilityUsers.filter((item,i)=>AvalabilityUsers.indexOf(item)==i);
// console.log(AvalabilityUsers);
let UserWithPercentage = []
AvalabilityUsers.forEach((users)=>{

  let userPercentage = 0;
  let UserLocation = "";
  let UserAffiliation = "";
  AllAvailabilityDetails.forEach((all)=>{
    all.UserName == users?userPercentage += parseInt(all.Percentage):userPercentage += 0;

  })
  UserDetails.forEach((UDetails)=>{
    UDetails.Name == users?UserLocation = UDetails.Location:"";
    UDetails.Name == users?UserAffiliation = UDetails.Affiliation:"";
  })
  UserWithPercentage.push({UserName:users,Percentage:userPercentage,Location:UserLocation,UserAff:UserAffiliation});
  // console.log({users:users,userPercentage:userPercentage,Location:UserLocation});
  
});


UserWithPercentage.forEach((avli)=>{
  AvailHtml+=`<tr><td>${avli.UserName}</td>
  <td>${avli.Location}</td>
  <td>${avli.UserAff}</td>
  <td>
  <div class="d-flex align-item-center">  
  <div class="availability-progress-bar" style="border: 1px solid  ${avli.Percentage >= 50? "#f01616":"#45b345"}">
  <div class="progress-value" style="height:100%;width:${100 - avli.Percentage}%; background: ${avli.Percentage >= 50? "#f01616":"#45b345"}"></div>
  </div>
  <span style="color:${avli.Percentage >= 50? "#f01616":"#45b345"}">${100 - avli.Percentage}%</span></div>
  </td></tr>`
})
$('#StaffAvailabilityTbody').html(AvailHtml)

  UserDetails.forEach((details) => {
    // console.log(details.PhoneNumber.split("^"));
    let ViewPhoneNumber = details.PhoneNumber.split("^"); 
    ViewPhoneNumber.pop();
   
    if (details.Affiliation == "Employee") {       
      EmpTable += `<tr><td class="user-details-td"><div  class="usernametag">${details.Name}</div><div class="HUserDetails">
      <img src="" class="userimg"/>
      <div class="user-name">${details.Name}</div>
      <div class="user-JTitle">${details.JobTitle}</div>
      <div class="user-location">${details.Location}</div>
      <div class="user-avail-title">Availability</div> 
      </div></td><td>${   
        details.FirstName
      }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
        ","
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
    if (details.Affiliation == "Outside Consultant") {
      OutTable += `<tr><td class="user-details-td"><div  class="usernametag">${details.Name}</div><div class="HUserDetails">
      <img src="" class="userimg"/>
      <div class="user-name">${details.Name}</div>
      <div class="user-JTitle">${details.JobTitle}</div>
      <div class="user-location">${details.Location}</div>
      <div class="user-avail-title">Availability</div> 
      </div></td><td>${
        details.FirstName
      }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
        ","
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
    if (details.Affiliation == "Affiliate") {
      AffTable += `<tr><td class="user-details-td"><div  class="usernametag">${details.Name}</div><div class="HUserDetails">
      <img src="" class="userimg"/>
      <div class="user-name">${details.Name}</div>
      <div class="user-JTitle">${details.JobTitle}</div>
      <div class="user-location">${details.Location}</div>
      <div class="user-avail-title">Availability</div> 
      </div></td><td>${
        details.FirstName
      }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
        ","
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
    if (details.Affiliation == "Alumni") {
      AlumTable += `<tr><td class="user-details-td"><div  class="usernametag">${details.Name}</div><div class="HUserDetails">
      <img src="" class="userimg"/>
      <div class="user-name">${details.Name}</div>
      <div class="user-JTitle">${details.JobTitle}</div>
      <div class="user-location">${details.Location}</div>
      <div class="user-avail-title">Availability</div> 
      </div></td><td>${
        details.FirstName
      }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
        ","
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
    AllDetailsTable += `<tr><td class="user-details-td"><div  class="usernametag">${details.Name}</div><div class="HUserDetails">
    <img src="" class="userimg"/>
    <div class="user-name">${details.Name}</div>
    <div class="user-JTitle">${details.JobTitle}</div>
    <div class="user-location">${details.Location}</div>
    <div class="user-avail-title">Availability</div> 
    </div></td><td>${
      details.FirstName
    }</td><td>${details.LastName}</td><td>${ViewPhoneNumber.join(
      ","
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
    BillingRateTable += `<tr><td class="usernametag">${details.Name}</td><td>${
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
      details.USDHourly == "" || details.USDHourly == null
        ? ""
        : `USD: ${details.USDHourly}`
    }</div><div>${
      details.EURHourly == "" || details.EURHourly == null
        ? ""
        : `EUR: ${details.EURHourly}`
    }</div><div>${
      details.OtherCurrHourly == "" || details.OtherCurrHourly == null
        ? ""
        : `${details.OtherCurr}: ${details.OtherCurrHourly}`
    }</div></td><td>${
      details.EffectiveDate == "Not Available" ? "Not Available" : new Date(details.EffectiveDate).toLocaleDateString()
    }</td></tr>`;
  });






  $("#SdhEmpTbody").html(EmpTable);
  $("#SdhOutsideTbody").html(OutTable);
  $("#SdhAffilateTbody").html(AffTable);
  $("#SdhAllumniTbody").html(AlumTable);
  $("#SdhAllPeopleTbody").html(AllDetailsTable);
  $("#SdgofficeinfoTbody").html(OfficeTable);
  $("#SdgBillingrateTbody").html(BillingRateTable);
  var options = {
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
  // bindStaffAvailTable(options);
  SdhEmpTableRowGrouping(1,"StaffAvailabilityTable",bindStaffAvailTable);
  UserProfileDetail();
     
}
const bindEmpTable = (options) => {
  (<any>$("#SdhEmpTable")).DataTable(options);
};
const bindOutTable = (options) => {
  (<any>$("#SdhOutsideTable")).DataTable(options);
};
const bindAffTable = (options) => {
  (<any>$("#SdhAffilateTable")).DataTable(options);
};
const bindAlumTable = (options) => {
  (<any>$("#SdhAllumniTable")).DataTable(options);
};
const bindAllDetailTable = (options) => {
  (<any>$("#SdhAllPeopleTable")).DataTable(options);
};
const bindOfficeTable = (option) => {
  (<any>$("#SdgofficeinfoTable")).DataTable(option);
};
const bindBillingRateTable = (option) => {
  (<any>$("#SdgBillingrateTable")).DataTable(option);
};
const bindStaffAvailTable = (option) =>{
  (<any>$('#StaffAvailabilityTable')).DataTable(option);
}
//Todo TableRowGrouping
const SdhEmpTableRowGrouping = (colno, tablename, tablefn) => {
  var collapsedGroups = {};
  var options = {
    order: [[colno, "asc"]],
    destroy: true,
    rowGroup: {
      // Uses the 'row group' plugin
      dataSrc: colno,
      startRender: function (rows, group) {
        var collapsed = !!collapsedGroups[group];
        rows.nodes().each(function (r) {
          r.style.display = collapsed ? "none" : "";
        });
        // Add category name to the <tr>. NOTE: Hardcoded colspan
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
    // table.draw(false);
  });
  tablefn(options);
  // UserProfileDetail();
};

function startIt() {
  var schema = {};
  schema["PrincipalAccountType"] = "User,DL,SecGroup,SPGroup";
  schema["SearchPrincipalSource"] = 15;
  schema["ResolvePrincipalSource"] = 15;
  schema["AllowMultipleValues"] = false;
  schema["MaximumEntitySuggestions"] = 50;
  schema["Width"] = "280px";

  // Render and initialize the picker.
  // Pass the ID of the DOM element that contains the picker, an array of initial
  // PickerEntity objects to set the picker value, and a schema that defines
  // picker properties.
  SPClientPeoplePicker_InitStandaloneControlWrapper(
    "peoplepickerText",
    null,
    schema
  );
}

const UserProfileDetail = async () => {


  //(<any>$("#peoplepickerText")).spPeoplePicker();

  const viewDir = document.querySelector(".view-directory");
  const editDir = document.querySelector(".edit-directory");
  const editbtn = document.querySelector(".edit-btn");
  const submitbtn = document.querySelector("#BtnSubmit");
  const cancelbtn = document.querySelector("#BtnCancel");
  ItemID = 0;
  OfficeAddArr = [];
  let office = await sp.web.getList(listUrl + "SDGOfficeInfo").items.get();
  office.forEach((off) => {
    OfficeAddArr.push({ OfficePlace: off.Office, OfficeFullAdd: off.Address });
  });
  SelectedUser = "";
  const userpage = document.querySelector(".user-profile-page");
  const username = document.querySelectorAll(".usernametag");
  const sdhEmp = document.querySelector(".sdh-employee");

  const Edit = document.querySelector("#btnEdit");
  const UserView = document.querySelector(".view-directory");
  const UserEdit = document.querySelector(".edit-directory");
  const StaffStatus = document.querySelector("#staffstatusDD");
  const WorkscheduleSection = document.querySelector("#workscheduleSec");
  SelectedUserProfile = [];
  username.forEach((btn) => {
    btn.addEventListener("click", async (e) => {
      if (!sdhEmp.classList.contains("hide")) {
        sdhEmp.classList.add("hide");
        userpage.classList.remove("hide");
      }
      SelectedUser = e.currentTarget["textContent"];
      
      SelectedUserProfile = UserDetails.filter((li) => {
        return li.Name == SelectedUser;
      });
      
      useravailabilityDetails(); 
      const specificUser = graph.users
        .getById(SelectedUserProfile[0].Usermail)
        .photo.getBlob()
        .then((photo: any) => {
          // console.log(photo);
          const url = window.URL;
          const blobUrl = url.createObjectURL(photo);
          $(".profile-picture").attr("src", blobUrl);
        })
        .catch((err) => {
          $(".profile-picture").attr(
            "src",
            "data:image/png;base64,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"
          );
        });
      let filesHtml = "";
      let editfileHtml = "";
      let files = await sp.web
        .getFolderByServerRelativeUrl(
          `BiographyDocument/${SelectedUserProfile[0].Usermail}`
        )
        .files.get();
      // console.log(files);
      files.forEach((file) => {
        // console.log(file.Name.split(".").pop());

        if (
          file.Name.split(".").pop() == "doc" || 
          file.Name.split(".").pop() == "docx"
        ) {
          filesHtml += `<div class="doc-section"><span class="word-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank">${file.Name}</a></div>`;
          editfileHtml += `<div class="quantityFiles"><span class="upload-filename">${file.Name}</span><a filename="${file.Name}" class="clsfileremove">x</a></div>`;
        } else if (file.Name.split(".").pop() == "xlsx" || file.Name.split(".").pop() == "csv") {
          filesHtml += `<div class="doc-section"><span class="excel-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank">${file.Name}</a></div>`;
          editfileHtml += `<div class="quantityFiles"><span class="upload-filename">${file.Name}</span><a  filename="${file.Name}" class="clsfileremove">x</a></div>`;
        } else if (
          file.Name.split(".").pop() == "png" ||
          file.Name.split(".").pop() == "jpg" ||
          file.Name.split(".").pop() == "jpeg"
        ) {
          filesHtml += `<div class="doc-section"><span class="pic-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank">${file.Name}</a></div>`;
          editfileHtml += `<div class="quantityFiles"><span class="upload-filename">${file.Name}</span><a  filename="${file.Name}" class="clsfileremove">x</a></div>`;
        } else {
          filesHtml += `<div class="doc-section"><span class="new-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank">${file.Name}</a></div>`;
          editfileHtml += `<div class="quantityFiles"><span class="upload-filename">${file.Name}</span><a  filename="${file.Name}" class="clsfileremove">x</a></div>`;
        }
      });
  
      let billingRateHtml = "";
      if (SelectedUserProfile[0].USDDaily != null && SelectedUserProfile[0].USDDaily != 0 && SelectedUserProfile[0].USDDaily != "0")
      {
       
        billingRateHtml += `<div class="billing-rates"><label>USD Daily Rates</label><div class="usd-daily-rate lblBlue" id="UsdDailyRate">${SelectedUserProfile[0].USDDaily}</div></div><div class="billing-rates"><label>USD Hourly Rates</label><div class="usd-hourly-rate lblBlue" id="UsdHourlyRate">${SelectedUserProfile[0].USDHourly}</div></div>`;
      }

      if (
        SelectedUserProfile[0].EURDaily != null &&
        SelectedUserProfile[0].EURDaily != 0 &&
        SelectedUserProfile[0].EURDaily != "0"
      ) {
        billingRateHtml += `<div class="billing-rates"><label>EUR Daily Rates</label><div class="eur-daily-rate lblBlue" id="EURDailyRate">${SelectedUserProfile[0].EURDaily}</div></div><div class="billing-rates"><label>EUR Hourly Rates</label><div class="eur-hourly-rate lblBlue" id="EURHourlyRate">${SelectedUserProfile[0].EURHourly}</div></div>`;
      }

      if (
        SelectedUserProfile[0].OtherCurrDaily != null &&
        SelectedUserProfile[0].OtherCurrDaily != 0 &&
        SelectedUserProfile[0].OtherCurrDaily != "0"
      ) {
        billingRateHtml += `<div class="billing-rates"><label>${SelectedUserProfile[0].OtherCurr} Daily Rates</label><div class="eur-daily-rate lblBlue" id="oDailyRate">${SelectedUserProfile[0].OtherCurrDaily}</div></div><div class="billing-rates"><label>${SelectedUserProfile[0].OtherCurr} Hourly Rates</label><div class="eur-hourly-rate lblBlue" id="oHourlyRate">${SelectedUserProfile[0].OtherCurrHourly}</div></div>`;
      }
      if (SelectedUserProfile[0].EffectiveDate != null) {
        billingRateHtml += ` <div class="billing-effective-date"><label>Effective Date</label><div class="effective-date lblBlue" id="EffectiveDate">${new Date(
          SelectedUserProfile[0].EffectiveDate
        ).toLocaleDateString()}</div></div>`;
      }
      ItemID = SelectedUserProfile[0].ItemID;
 
      selectedUsermail = SelectedUserProfile[0].Usermail;
      // console.log(selectedUsermail);
      if(SelectedUserProfile[0].UserPersonalMail != "" && SelectedUserProfile[0].UserPersonalMail != null) {

        $('#userpersonalmail').parent().removeClass('hide');
         $('#userpersonalmail').html(SelectedUserProfile[0].UserPersonalMail)
      } 
      else{
        $('#userpersonalmail').html("")
        $('#userpersonalmail').parent().addClass('hide');
      }
      if(SelectedUserProfile[0].Assistant != null && SelectedUserProfile[0].Assistant !=""){
        $("#viewAssistant").html(`<div class="d-flex align-item-center">
        <label>Assistant : </label><div class="lblRight" id="assistantViewpage">${SelectedUserProfile[0].Assistant}</div>
        </div>`)
      }else{
        $("#viewAssistant").html("")  
      } 
      
      if(SelectedUserProfile[0].PhoneNumber)
          {
              var html="";
              var phno=SelectedUserProfile[0].PhoneNumber
              var val=phno.split("^");
              console.log("Split");
              console.log(val);
              if(val.length>1)
              {
                for(var i=0;i<val.length-1;i++)
              {
                var temp=val[i].split("-");
                if(temp[1]==" ")
                {
                  html+="Not Available";
                }
                else{
                  html+=val[i]+";";
                }
              }
              $("#user-phone").html(html);
              }
              else{
                $("#user-phone").val("Not Available");
              }
            }
      // $('#linkedinIDview').html(`<span class="linkedInBtn" id="linkedInBtn">Link</span>`);
      $('#linkedinIDview').html(`<a href="${SelectedUserProfile[0].LinkedInID.Url}" target ='_blank' data-interception="off"><span class="icon-linkedin"></span></a>`);
      $('#PSignOther').html(SelectedUserProfile[0].SignOther);
      $('#PChildren').html(SelectedUserProfile[0].Child);
      $("#user-Designation").html(SelectedUserProfile[0].Affiliation);
      $("#user-staff-function").html(SelectedUserProfile[0].Title)
      $("#user-job-title").html(SelectedUserProfile[0].JobTitle);
      $("#user-location").html(SelectedUserProfile[0].Location);
      // $("#user-office").html(SelectedUserProfile[0].Location);
      // $("#user-phone").html(SelectedUserProfile[0].PhoneNumber);
      // $("#user-mail").html(SelectedUserProfile[0].Usermail);
      // $("#personal-mail").html(SelectedUserProfile[0].UserPersonalMail);
      $("#UserProfileName").html(SelectedUserProfile[0].Name);
      $("#UserProfileEmail").html(
        `<span class="user-mail-icon"></span>${SelectedUserProfile[0].Usermail}`
      );
      $("#PAddLine").html(SelectedUserProfile[0].HAddLine);
      $("#PAddCity").html(SelectedUserProfile[0].HAddCity);
      $("#PAddState").html(SelectedUserProfile[0].HAddState);
      $("#PAddPCode").html(SelectedUserProfile[0].HAddPCode);
      $("#PAddPCountry").html(SelectedUserProfile[0].HAddPCountry);
      $("#WAddressDetails").html(
        OfficeAddArr.filter(
          (add) => SelectedUserProfile[0].Location == add.OfficePlace
        )[0].OfficeFullAdd
      );
      $("#WLoctionDetails").html(SelectedUserProfile[0].Location)
      $("#shortbio").html(SelectedUserProfile[0].ShortBio);
      $("#citizenship").html(SelectedUserProfile[0].Citizen);
      $("#IndustryExp").html(SelectedUserProfile[0].Industry);
      $("#LanguageExp").html(SelectedUserProfile[0].Language);
      $("#SDGCourse").html(SelectedUserProfile[0].SDGCourse);
      $("#SoftwareExp").html(SelectedUserProfile[0].Software);
      $("#MembershipExp").html(SelectedUserProfile[0].Membership);
      $("#SpecialKnowledge").html(SelectedUserProfile[0].SpecialKnowledge);
      $("#BillingRateDetails").html(billingRateHtml);
      $("#bioAttachment").html(filesHtml);
      $("#filesfromfolder").html(editfileHtml);
      $("#staffStatus").html(SelectedUserProfile[0].StaffStatus);
      $("#workscheduleViewSec").html(
        SelectedUserProfile[0].StaffStatus == "Part-time"
          ? `
      <div class="d-flex"><label>Work Schedule</label><p class="lblRight" id="workSchedule">${SelectedUserProfile[0].WorkSchedule == null ||SelectedUserProfile[0].WorkSchedule ==  ""?"Not Available":SelectedUserProfile[0].WorkSchedule}</p></div>`
          : ""
      );
      var finalmonth: any = "";
      var dd = new Date(SelectedUserProfile[0].EffectiveDate).getDate();
      var mm = new Date(SelectedUserProfile[0].EffectiveDate).getMonth() + 1;
      mm < 10 ? (finalmonth = "0" + mm) : (finalmonth = mm);
      var yyyy = new Date(SelectedUserProfile[0].EffectiveDate).getFullYear();
      var dateformat = yyyy + "-" + finalmonth + "-" + dd;
      $("#EffectiveDateEdit").val(dateformat);
      
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
    sp.web
      .getFileByServerRelativeUrl(
        `/sites/StaffDirectory/BiographyDocument/${SelectedUserProfile[0].Usermail}/${filename}`
      )
      .recycle()
      .then(function (data) {});
  });
  
};

const editFunction = async() => {
  await SPComponentLoader.loadScript("/_layouts/15/init.js").then(() => {});
  await SPComponentLoader.loadScript("/_layouts/15/1033/sts_strings.js");
  await SPComponentLoader.loadScript("/_layouts/15/clientforms.js");
  await SPComponentLoader.loadScript("/_layouts/15/clienttemplates.js");
  await SPComponentLoader.loadScript("/_layouts/15/clientpeoplepicker.js");
  await SPComponentLoader.loadScript("/_layouts/15/autofill.js");
  await SPComponentLoader.loadScript("/_layouts/15/SP.js");
  await SPComponentLoader.loadScript("/_layouts/15/sp.runtime.js");
  await SPComponentLoader.loadScript("/_layouts/15/sp.core.js");

  await startIt();
  const Edit = document.querySelector("#btnEdit");
  const UserView = document.querySelector(".view-directory");
  const UserEdit = document.querySelector(".edit-directory");

  if (!UserView.classList.contains("hide")) {
    UserView.classList.add("hide");
    UserEdit.classList.remove("hide");
    Edit.classList.add("hide");
  } else {
    UserEdit.classList.remove("hide");
    Edit.classList.add("hide");
  }

  let MobileNumberHtmlSec = "";
  let HomeNumberHtmlSec = "";
  let EmergencyNumberHtmlSec = "";
  

  let MCCodeArr = []
  if (
    SelectedUserProfile[0].PhoneNumber != "" && SelectedUserProfile[0].PhoneNumber != null
  ) {
    let AllMnumber = SelectedUserProfile[0].PhoneNumber.split("^");
    AllMnumber.pop();
    let AllMobileNumbers = AllMnumber;
    
    AllMobileNumbers.forEach((numbers, i) => {
      // console.log(CCodeHtml)
      let SplitedMNum = numbers.split(" - ");
      MCCodeArr.push(SplitedMNum[0])

      if (i == 0) {
        MobileNumberHtmlSec += `<div class="d-flex mobNumbers"><select class="mobNoCode">${CCodeHtml}</select><input type="number" class="mobNo" id="" value="${SplitedMNum[1]}"><span class="addMobNo add-icon"></div>`;
      } else {
        MobileNumberHtmlSec += `<div class="d-flex mobNumbers"><select class="mobNoCode">${CCodeHtml}</select><input type="number" class="mobNo" id="" value="${SplitedMNum[1]}"><span class="removeMobNo remove-icon"></div>`;
      }
      $("#mobileNoSec").html(MobileNumberHtmlSec);
    });

  } else {
    MobileNumberHtmlSec += `<div class="d-flex mobNumbers"><select class="mobNoCode">${CCodeHtml}</select><input type="number" class="mobNo" id=""><span class="addMobNo add-icon"></div>`;
    $("#mobileNoSec").html(MobileNumberHtmlSec);
  }
  let HCCodeArr=[];
if(SelectedUserProfile[0].HomeNo != "" && SelectedUserProfile[0].HomeNo != null){
  let AllHNumber = SelectedUserProfile[0].HomeNo.split("^");
  AllHNumber.pop();
  let AllHomeNumber = AllHNumber;
  AllHomeNumber.forEach((hnumbs,j)=>{
    let SplitedHNum = hnumbs.split(' - ');
    HCCodeArr.push(SplitedHNum[0])
    if(j == 0){
      HomeNumberHtmlSec +=`<div class="d-flex homeNumbers"><select class="homeNoCode">${CCodeHtml}</select><input type="number" class="home" id="" value="${SplitedHNum[1]}"><span class="addHomeNo add-icon"></div>`
    }else{
      HomeNumberHtmlSec +=`<div class="d-flex homeNumbers"><select class="homeNoCode">${CCodeHtml}</select><input type="number" class="home" id="" value="${SplitedHNum[1]}"><span class="removeHomeNo remove-icon"></div>`
    }
  });
  $('#homeNoSec').html(HomeNumberHtmlSec);
}
else{
  HomeNumberHtmlSec +=`<div class="d-flex homeNumbers"><select class="homeNoCode">${CCodeHtml}</select><input type="number" class="home" id=""><span class="addHomeNo add-icon"></div>`
  $('#homeNoSec').html(HomeNumberHtmlSec);
}

let ECCodeArr = [];
if(SelectedUserProfile[0].EmergencyNo != ""  && SelectedUserProfile[0].EmergencyNo != null){
  let AllENumber = SelectedUserProfile[0].EmergencyNo.split("^");
  AllENumber.pop();
  let AllEmergencyNumber = AllENumber;
  AllEmergencyNumber.forEach((enums,k)=>{
    let SplitedENum = enums.split(' - ');
    ECCodeArr.push(SplitedENum[0])
    if(k==0){
      EmergencyNumberHtmlSec +=`<div class="d-flex emergencyNumbers"><select class="emergencyNoCode" id="ec${k}">${CCodeHtml}</select><input type="number" class="home" id="" value="${SplitedENum[1]}"><span class="addEmergencyNo add-icon"></div>`
    }else{
      EmergencyNumberHtmlSec +=`<div class="d-flex emergencyNumbers"><select class="emergencyNoCode" id="ec${k}">${CCodeHtml}</select><input type="number" class="home" id="" value="${SplitedENum[1]}"><span class="removeEmergencyNo remove-icon"></div>`
    }
  })
  $('#emergencyNoSec').html(EmergencyNumberHtmlSec);
}else{
  EmergencyNumberHtmlSec +=`<div class="d-flex emergencyNumbers"><select class="emergencyNoCode">${CCodeHtml}</select><input type="number" class="home" id=""><span class="addEmergencyNo add-icon"></div>`
  $('#emergencyNoSec').html(EmergencyNumberHtmlSec);
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

  $("#EditedAddressDetails").html(OfficeAddArr.filter(
    (add) => SelectedUserProfile[0].Location == add.OfficePlace
  )[0].OfficeFullAdd);
  $("#StaffFunctionEdit").val(SelectedUserProfile[0].Title);
  $("#StaffAffiliatesEdit").val(SelectedUserProfile[0].Affiliation);
  $("#PAddLineE").val(SelectedUserProfile[0].HAddLine);
  $("#PAddCityE").val(SelectedUserProfile[0].HAddCity);
  $("#PAddStateE").val(SelectedUserProfile[0].HAddState);
  $("#PAddPCodeE").val(SelectedUserProfile[0].HAddPCode);
  $("#PAddCountryE").val(SelectedUserProfile[0].HAddPCountry);
  $("#Eshortbio").val(SelectedUserProfile[0].ShortBio);
  $("#EIndustry").val(SelectedUserProfile[0].Industry);
  $("#ELanguage").val(SelectedUserProfile[0].Language);
  $("#ESDGCourse").val(SelectedUserProfile[0].SDGCourse);
  $("#ESoftwarExp").val(SelectedUserProfile[0].Software);
  $("#EMembership").val(SelectedUserProfile[0].Membership);
  $("#ESKnowledge").val(SelectedUserProfile[0].SpecialKnowledge);
  $("#citizenshipE").val(SelectedUserProfile[0].Citizen);
  $("#linkedInID").val(SelectedUserProfile[0].LinkedInID.Url);
  // $("#mobileno").val(SelectedUserProfile[0].PhoneNumber);
  $("#children").val(SelectedUserProfile[0].Child);
  $("#significantOther").val(SelectedUserProfile[0].SignOther);
  $("#USDDailyEdit").val(SelectedUserProfile[0].USDDaily);
  $("#USDHourlyEdit").val(SelectedUserProfile[0].USDHourly);
  $("#EURDailyEdit").val(SelectedUserProfile[0].EURDaily);
  $("#EURHourlyEdit").val(SelectedUserProfile[0].EURHourly);
  $("#assisstantName").val(SelectedUserProfile[0].Assistant);
  $("#personalmailID").val(SelectedUserProfile[0].UserPersonalMail);
  // $("#homeno").val(SelectedUserProfile[0].HomeNo);
  // $("#emergencyno").val(SelectedUserProfile[0].EmergencyNo);
  $("#workLocationDD").val(SelectedUserProfile[0].Location);
  $("#staffstatusDD").val(SelectedUserProfile[0].StaffStatus);
  $("#othercurrDD").val(SelectedUserProfile[0].OtherCurr);
  $("#ODailyEdit").val(SelectedUserProfile[0].OtherCurrDaily);
  $("#OHourlyEdit").val(SelectedUserProfile[0].OtherCurrHourly);

  if(SelectedUserProfile[0].StaffStatus == "Part-time"){
    $("#workscheduleEdit").html("");
    $("#workscheduleEdit").html(`<div class="d-flex w-100" id="workscheduleSec"> <label>Work Schedule</label><div class="w-100"><input type="text" id="workScheduleE" value="${SelectedUserProfile[0].WorkSchedule == "" || SelectedUserProfile[0].WorkSchedule == null ? "":SelectedUserProfile[0].WorkSchedule}"></div></div>`)
  }
  else{
    $("#workscheduleEdit").html("");
    $("#workscheduleEdit").html(`<div class="d-flex w-100 hide" id="workscheduleSec"> <label>Work Schedule</label><div class="w-100"><input type="text" id="workScheduleE" value="${SelectedUserProfile[0].WorkSchedule == "" || SelectedUserProfile[0].WorkSchedule == null ? "":SelectedUserProfile[0].WorkSchedule}"></div></div>`)
  }
  
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
  // setTimeout(() =>{
  //   console.log($("#peoplepickerText"));
  // },10000)
  
  // $('.sp-peoplepicker-userDisplayLink').text();
};
const editsubmitFunction = async () => {
  let mobNumUpdate = "";
  let homeNumUpdate = "";
  let emergencyNumUpdate = "";
  let mobNumbers = document.querySelectorAll(".mobNumbers");
  let homeNumbers = document.querySelectorAll(".homeNumbers");
  let emergencyNumbers = document.querySelectorAll(".emergencyNumbers");
  mobNumbers.forEach((nums) => {
    mobNumUpdate += `${CCodeArr[nums.children[0]["options"].selectedIndex]} - ${
      nums.children[1]["value"]
    }^`;
  });
  homeNumbers.forEach((nums) => {
    homeNumUpdate += `${
      CCodeArr[nums.children[0]["options"].selectedIndex]
    } - ${nums.children[1]["value"]}^`;
  });
  emergencyNumbers.forEach((nums) => {
    emergencyNumUpdate += `${
      CCodeArr[nums.children[0]["options"].selectedIndex]
    } - ${nums.children[1]["value"]}^`;
  });

  if (bioAttachArr.length > 0) {
    bioAttachArr.map((filedata) => {
      sp.web.folders
        .add(`/sites/StaffDirectory/BiographyDocument/${selectedUsermail}`)
        .then((data) => {
          sp.web
            .getFolderByServerRelativeUrl(data.data.ServerRelativeUrl)
            .files.add(filedata.name, filedata.content, true);
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
  

  try {
    const update = await sp.web
    .getList(listUrl + "StaffDirectory")
    .items.getById(ItemID)
    .update({
      Title: "SDG User Info",
      PersonalEmail: $("#personalmailID").val(),
      MobileNo: mobNumUpdate,
      HomeNo: homeNumUpdate,
      EmergencyNo: emergencyNumUpdate,
      HomeAddLine: $("#PAddLineE").val(),
      HomeAddCity: $("#PAddCityE").val(),
      HomeAddState: $("#PAddStateE").val(),
      HomeAddPCode: $("#PAddPCodeE").val(),
      HomeAddCountry: $("#PAddCountryE").val(),
      IndustryExp: $("#EIndustry").val(),
      LanguageExp: $("#ELanguage").val(),
      SDGCourses: $("#ESDGCourse").val(),
      SoftwareExp: $("#ESoftwarExp").val(),
      Membership: $("#EMembership").val(),
      SpecialKnowledge: $("#ESKnowledge").val(),
      Citizenship: $("#citizenshipE").val(),
      ShortBio: $("#Eshortbio").val(),
      USDDailyRate: $("#USDDailyEdit").val(),
      USDHourlyRate: $("#USDHourlyEdit").val(),
      EURDailyRate: $("#EURDailyEdit").val(),
      EURHourlyRate: $("#EURHourlyEdit").val(),
      OtherCurrency: $("#othercurrDD").val(),
      ODailyRate: $("#ODailyEdit").val(),
      OHourlyRate: $("#OHourlyEdit").val(),
      EffectiveDate: $("#EffectiveDateEdit").val(),
      signother: $("#significantOther").val(),
      children: $("#children").val(),
      WorkingSchedule: $("#workSchedule").val(),
      SDGOffice: $("#workLocationDD").val(),
      StaffStatus: $("#staffstatusDD").val(),
      LinkedInLink:{
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: "LinkedIn",
        Url: $("#linkedInID").val()
   
      },
      stafffunction:$("#StaffFunctionEdit").val(),
      SDGAffiliation:$("#StaffAffiliatesEdit").val(),
      AssistantId: profileID,
    });
  location.reload();
  } catch (error) {
    console.log(error);
    
    
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

const useravailabilityDetails = async() =>{
  console.log(SelectedUserProfile[0].Name);
  
  let availList = await sp.web
  .getList(listUrl + "SDGAvailability")
  .items.select("*","UserName").filter(`UserName eq '${SelectedUserProfile[0].Name}'`)
  .getAll();
  console.log(availList); 
  
  let availTableHtml = "";
  availList.forEach((avail)=>{
    availTableHtml  += `<tr><td>${avail.Project}</td><td>${new Date(avail.StartDate).toLocaleDateString()}</td><td>${new Date(avail.EndDate).toLocaleDateString()}</td><td>${avail.Percentage}%</td><td>${avail.Comments}</td><td><div class="d-flex"><div class="action-btn action-edit" data-toggle="modal" data-target="#addprojectmodal" data-id="${avail.ID}" id="editProjectAvailability"></div><div class="action-btn action-delete" data-id="${avail.ID}" id="deleteProjectAvailability"> </div></div></td></tr>`
  });
  
  $("#UserAvailabilityTbody").html("");
  $("#UserAvailabilityTbody").html(availTableHtml);
  var options = {  
    order: [[0, "asc"]],
    destroy:true,
    
  }; 
 var userAvailTable =  (<any>$("#UserAvailabilityTable")).DataTable(options);
 $('.usernametag').on( 'click', function () {
  userAvailTable.destroy();
} );
}

function removeSelectedfile(filename) {
  for (var i = 0; i < bioAttachArr.length; i++) {
    if (bioAttachArr[i].name == filename) {
      ///filesQuantity[i].remove();
      bioAttachArr.splice(i, 1);
      break;
    } 
  }
}
const removeAvailProject = (ID) =>{
  sp.web.getList(listUrl + "SDGAvailability").items.getById(ID).delete();
}
const multipleMobNo = () => {
  $("#mobileNoSec").append(  
    `<div class="d-flex mobNumbers"><select class="mobNoCode">${CCodeHtml}</select><input type="number" class="mobNo" id="mobileno1"/><span class="removeMobNo remove-icon"></span></div>`
  );
};
const multipleHomeNo = () => {
  $("#homeNoSec").append(
    `<div class="d-flex homeNumbers"><select class="homeNoCode">${CCodeHtml}</select><input type="number" class="mobNo" id="homeno1"/><span class="removeHomeNo remove-icon"></span></div>`
  );
};
const multipleEmergencyNo = () => {
  $("#emergencyNoSec").append(
    `<div class="d-flex emergencyNumbers"><select class="emergencyNoCode">${CCodeHtml}</select><input type="number" class="mobNo" id="emergencyno1"/><span class="removeHomeNo remove-icon"></span></div>`
  );
};
const fillEditSection = (ID) =>{
  
   sp.web
  .getList(listUrl + "SDGAvailability").items.getById(ID).get().then((item:any)=>{
    var Sfinalmonth: any = "";
      var Sdd = new Date(item.StartDate).getDate();
      var Smm = new Date(item.StartDate).getMonth() + 1;
      Smm < 10 ? (Sfinalmonth = "0" + Smm) : (Sfinalmonth = Smm);
      var Syyyy = new Date(item.StartDate).getFullYear();
      var Sdateformat = Syyyy + "-" + Sfinalmonth + "-" + Sdd;

      var Efinalmonth: any = "";
      var Edd = new Date(item.EndDate).getDate();
      var Emm = new Date(item.EndDate).getMonth() + 1;
      Emm < 10 ? (Efinalmonth = "0" + Smm) : (Efinalmonth = Emm);
      var Eyyyy = new Date(item.EndDate).getFullYear();
      var Edateformat = Eyyyy + "-" + Efinalmonth + "-" + Edd;
    $("#projectName").val(item.Project);
    $("#projectStartDate").val(Sdateformat);
    $("#projectEndDate").val(Edateformat);
    $("#projectPercent").val(item.Percentage);
    $("#practiceAreaDD").val(item.ProjectArea);
    $("#client").val(item.Client);
    $("#projectCode").val(item.ProjectCode);
    $("#ProjectLocation").val(item.ProjectLocation);
    $("#projectAvailNotes").val(item.Notes);
    $("#Projectcomments").val(item.Comments)
  })
}
const availSubmitFunc = async() =>{

  let availdetails = await sp.web
  .getList(listUrl + "SDGAvailability")
  .items.select("*","UserName").filter(`UserName eq '${SelectedUserProfile[0].Name}'`)
  .getAll();
  console.log("availdetails");
  console.log(availdetails); 
  let strdate=$("#projectStartDate").val();
  let enddate=$("#projectEndDate").val();


  let ProjectPercent = 0;
  $("#projectPercent").val() == ""?ProjectPercent=0:ProjectPercent=parseInt(<any>$("#projectPercent").val())

  const submitProject = await sp.web
  .getList(listUrl + "SDGAvailability")
  .items.add(
    {
      UserName:SelectedUserProfile[0].Name,
      Project:$("#projectName").val(),
      StartDate:$("#projectStartDate").val(),
      EndDate:$("#projectEndDate").val(),
      Percentage:ProjectPercent.toString(),
      ProjectArea:$("#practiceAreaDD").val(),
      Client:$("#client").val(),
      ProjectCode:$("#projectCode").val(),
      ProjectLocation:$("#ProjectLocation").val(),
      Notes:$("#projectAvailNotes").val(),
      Comments:$("#Projectcomments").val()
    }
  )
  console.log(submitProject);
  $("#projectName").val("")
$("#projectStartDate").val("")
$("#projectEndDate").val("")
$("#projectPercent").val("")
$("#practiceAreaDD").val("")
$("#client").val("")
$("#projectCode").val("")
$("#ProjectLocation").val("")
$("#projectAvailNotes").val("")
$("#Projectcomments").val("")
location.reload();
}
const availUpdateFunc = () =>{
  let ProjectPercent = 0;
  $("#projectPercent").val() == ""?ProjectPercent=0:ProjectPercent=parseInt(<any>$("#projectPercent").val())
  const updateProject =  sp.web
  .getList(listUrl + "SDGAvailability").items.getById(AvailEditID).update(
    {
      UserName:SelectedUserProfile[0].Name,
      Project:$("#projectName").val(),
      StartDate:$("#projectStartDate").val(),
      EndDate:$("#projectEndDate").val(),
      Percentage:ProjectPercent.toString(),
      ProjectArea:$("#practiceAreaDD").val(),
      Client:$("#client").val(),
      ProjectCode:$("#projectCode").val(),
      ProjectLocation:$("#ProjectLocation").val(),
      Notes:$("#projectAvailNotes").val(),
      Comments:$("#Projectcomments").val()
    }
  );
  $("#projectName").val("")
  $("#projectStartDate").val("")
  $("#projectEndDate").val("")
  $("#projectPercent").val("")
  $("#practiceAreaDD").val("")
  $("#client").val("")
  $("#projectCode").val("")
  $("#ProjectLocation").val("")
  $("#projectAvailNotes").val("")
  $("#Projectcomments").val("")
  AvailEditID = 0;
  AvailEditFlag = false;
  location.reload();
}