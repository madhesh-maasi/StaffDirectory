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
import { graph } from "@pnp/graph";

import "@pnp/graph/users";
import "@pnp/graph/photos";
import "@pnp/sp/profiles"; 
// SPComponentLoader.loadScript(
//   "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.4.min.js"
//   // "https://code.jquery.com/jquery-3.5.1.js"
// );

import "../../ExternalRef/CSS/style.css";

SPComponentLoader.loadCss(
  "https://cdn.jsdelivr.net/npm/select2@4.1.0-beta.1/dist/css/select2.min.css"
);
SPComponentLoader.loadCss(
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"
);
import "datatables";

require("datatables.net-dt");
require("datatables.net-rowgroup-dt");
SPComponentLoader.loadCss(
  "https://cdn.datatables.net/1.10.24/css/jquery.dataTables.css"
);

var that:any;
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
         <li><a href="#">By Last Name</a></li>
         <li><a href="#">By First Name</a></li>
         </ul>
         </div>
         </div>
       </div>
     </div>
     <div class="card">
       <div class="card-header nav-items" id="headingSeven">
           <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseSeven" aria-expanded="false" aria-controls="collapseSeven"><span class="nav-icon staff-avail"></span>Staff Availability</div>
       </div>
       <div id="collapseSeven" class="clsCollapse collapse" aria-labelledby="headingSeven" data-parent="#accordionExample">
         <div class="card-body">
         <div class="filter-section">
         <ul>
         <li><a href="#">By Last Name</a></li>
         <li><a href="#">By First Name</a></li>
         </ul>
         </div>
         </div>
       </div>
     </div> 
     <div class="card">  
       <div class="card-header nav-items" id="headingEight">
           <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseEight" aria-expanded="false" aria-controls="collapseEight"><span class="nav-icon billing-rate"></span>Billing Rates</div>
       </div>
       <div id="collapseEight" class="clsCollapse collapse" aria-labelledby="headingEight" data-parent="#accordionExample">
         <div class="card-body">
         <div class="filter-section">
         <ul>
         <li><a href="#">By Last Name</a></li>
         <li><a href="#">By First Name</a></li>
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
     <div class="title-font" id="user-job-title">IT Manager</div>
     </div>
     <div class="user-info">   
     <label>Designation:</label> 
     <div class="title-font" id="user-Designation">Employee</div>
     </div>
     <div class="user-info">
     <label>Office:</label>
     <div class="title-font" id="user-office">Chennai</div>
     </div>
     </div>
     <div class="profile-details-right">
     <div class="user-info"> 
     <label>Mobile:</label>
     <div class="title-font" id="user-phone">+91 909090909</div>
     </div>
     <div class="user-info">   
     <label>Email:</label> 
     <div class="title-font" ><div id="user-mail"></div><div id="personal-mail"></div></div> 
     </div>
     </div> 
     </div>
     </div>
     <div class="user-profile-tabs">
     <div class="tab-section"> 
     <div class="tab-header-section"> 
     <ul class="nav nav-tabs">
       <li class="active"><a data-toggle="tab" href="#home">Directory Information</a></li>
       <li><a data-toggle="tab" href="#menu1">Availability</a></li>
     </ul>
     <div class="edit-btn"id="btnEdit"></div>
     </div>
     <div>
     <div class="tab-content">
    <div id="home" class="tab-pane fade in active">
      <div id="DirectoryInformation" class="d-flex view-directory">
      <div class="DInfo-left col-6">
      <div class="work-address">
      <h4>Work Address</h4>
      <div class="address-details lblRight" id="WAddressDetails">
  
      </div>
      </div>
      <div class="personal-info">
      <h4>Personal Info</h4>
      <div class="address-details" id="PersonaInfo"> 
      <div><label>Address Line:</label><label id="PAddLine" class="lblRight"></label></div>
      <div><label>City:</label><label id="PAddCity" class="lblRight"></label></div>
      <div><label>State:</label><label id="PAddState" class="lblRight"></label></div>
      <div><label>Postal Code :</label><label id="PAddPCode" class="lblRight"></label></div> 
      <div><label>Country:</label><label id="PAddPCountry" class="lblRight"></label></div>
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
      <div class="DInfo-right col-6">
      <div class="StaffStatus">
      <h4>Staff Status</h4>
      <p class="lblRight" id="staffStatus"></p>
      <div id="workscheduleViewSec">
      <h5>Work Schedule</h5>
      <p class="lblRight" id="workSchedule"></p>
      </div>
      </div>
      <div class="citizen-info">
      <h4>Citizenship</h4>
      <div class="address-details" id="CitizenInfo"> 
      <div><label>Citizen:</label><label id="citizenship" class="lblRight"></label></div>
      </div>
      </div>
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
      <div class="d-flex"><label>Mobile No :</label><div class="w-100"><textarea id="mobileno"></textarea></div></div>
      <div class="d-flex"><label>Home No :</label><div class="w-100"><textarea id="homeno"></textarea></div></div>
      <div class="d-flex"><label>Emergency No :</label><div class="w-100"><textarea id="emergencyno"></textarea></div></div>
      
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

  
      <div class="DInfo-right col-6">
      <div class="StaffStatus">
      <h4>Staff Status</h4> 
      <div class="d-flex w-100"> 
      <label>Status</label><div class="w-100"><select id="staffstatusDD"></select></div></div>
      <div class="d-flex w-100 hide" id="workscheduleSec">
      <label>Work Schedule</label>
      <div class="w-100"><input type="text" id="workScheduleE"></div>
      </div>
      </div> 
      <div class="citizen-info">
      <div class="address-details" id="CitizenInfo"> 
      <div class="d-flex w-100"><label>Citizen:</label><div class="w-100"><input type="text" id="citizenshipE"></div></div>
      </div>
      </div>
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
      </div>
      
      </div>
      <div class="btn-section">
      <button class="btn btn-cancel" id="BtnCancel">Cancel</button>
      <button class="btn btn-submit" id="BtnSubmit">Submit</button>
      </div>
      </div>
    </div>
    <div id="menu1" class="tab-pane fade">
      <h3>Availabilty Comes Here</h3>
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
    const editbtn = document.querySelector(".edit-btn");
    
    $(".clsToggleCollapse").click(function () {
      $(".clsCollapse").each(function () {
        $(this).removeClass("in").attr("style", "");
      });
      $(this).next("div").addClass("in");
    });
    onLoadData();
    ActiveSwitch();
    $(".SDHEmployee").click(() => { 
      if(viewDir.classList['contains']("hide")){
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
      if(viewDir.classList['contains']("hide")){
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
      if(viewDir.classList['contains']("hide")){
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
      if(viewDir.classList['contains']("hide")){
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
      if(viewDir.classList['contains']("hide")){
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
      if(viewDir.classList['contains']("hide")){
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
      }else{
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindOfficeTable(options);
      }
    });
    // Employee Filters
    $(".sdhLocgrouping").click(() => {
      SdhEmpTableRowGrouping(4, "SdhEmpTable", bindEmpTable);
    });
    $(".sdhTitlgrouping").click(() => {
      SdhEmpTableRowGrouping(6, "SdhEmpTable", bindEmpTable);
    });
    $(".sdhAssistantgrouping").click(() => {
      SdhEmpTableRowGrouping(7, "SdhEmpTable", bindEmpTable);
    });
    $(".sdhfirstnamesort").click(() => {  
      var options = {
        destroy: true,
        order: [[1, "asc"]],
      };
      bindEmpTable(options);
    });
    $(".sdhlastnamesort").click(() => {
      var options = {
        destroy: true,
        order: [[2, "asc"]],
      };
      bindEmpTable(options);
    });
    //OutSideConsultant
    $(".OutConslastnamesort").click(() => {
      var options = {
        destroy: true,
        order: [[2, "asc"]],
      };
      bindOutTable(options);
    });
    $(".OutConsFirstnamesort").click(() => {
      var options = {
        destroy: true,
        order: [[1, "asc"]],
      };
      bindOutTable(options);
    });
    $(".OutConsLocgrouping").click(() => {
      SdhEmpTableRowGrouping(4, "SdhOutsideTable", bindOutTable);
    });
    $(".OutConsStaffgrouping").click(() => {
      SdhEmpTableRowGrouping(6, "SdhOutsideTable", bindOutTable);
    });
    // Affliates
    $(".Afflastnamesort").click(() => {
      var options = {
        destroy: true,
        order: [[2, "asc"]],
      };
      bindAffTable(options);
    });
    $(".AffFirstnamesort").click(() => {
      var options = {
        destroy: true,
        order: [[1, "asc"]],
      };
      bindAffTable(options);
    });
    // Allumni
    $(".SDHAlumniLastName").click(() => {
      var options = {
        destroy: true,
        order: [[1, "asc"]],
      };
      bindAlumTable(options);
    });
    $(".SDHAlumniFirstName").click(() => {
      var options = {
        destroy: true,
        order: [[2, "asc"]],
      };
      bindAlumTable(options);
    });
    $(".SDHAlumniOffice").click(() => {
      SdhEmpTableRowGrouping(4, "SdhAllumniTable", bindAlumTable);
    });
    // All Users
    $(".SDHShowAllLastName").click(() => {
      var options = {
        destroy: true,
        order: [[2, "asc"]],
      };
      bindAllDetailTable(options);
    });
    $(".SDHShowAllFirstName").click(() => {
      var options = {
        destroy: true,
        order: [[1, "asc"]],
      };
      bindAllDetailTable(options);
    });





    $(document).on("change", "#BioAttachEdit", function () {
      if ($(this)[0].files.length > 0) {
        for (let index = 0; index < $(this)[0].files.length; index++) {
          const file = $("#BioAttachEdit")[0]['files'][index];
          // if (ValidateSingleInput($("#others")[0])) {
            bioAttachArr.push(file);
            $("#otherAttachmentFiles").append(
              '<div class="quantityFiles">' +
                "<span class=upload-filename>" +
                file.name +
                "</span>" +
                "<a filename='" +
                file.name +
                "' class=clsothersRemove href='#'>x</a></div>"
            );
          // }
        }
        $(this).val("");
        $(this).parent().find("label").text("Choose File");
      }
      console.log(bioAttachArr);
      
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
  let StaffStatusDD = document.querySelector("#staffstatusDD");
  let LocationDD = document.querySelector("#workLocationDD");
  let othercurrDD = document.querySelector("#othercurrDD");
  let LocOptionHtml ="";
  let staffOptionHtml = "";
  let otherCurrHtml ="";
  let listLocation = await sp.web.getList(listUrl + "StaffDirectory").fields.filter("EntityPropertyName eq 'SDGOffice'").get();
  let listStaffStatus = await sp.web.getList(listUrl + "StaffDirectory").fields.filter("EntityPropertyName eq 'StaffStatus'").get();
  let listOtherCurr = await sp.web.getList(listUrl + "StaffDirectory").fields.filter("EntityPropertyName eq 'OtherCurrency'").get();
 
  listLocation[0]['Choices'].forEach((li)=>{
    LocOptionHtml += `<option value="${li}">${li}</option>`
  })
  listStaffStatus[0]['Choices'].forEach((stff)=>{
    staffOptionHtml +=`<option value="${stff}">${stff}</option>`
  })
  listOtherCurr[0]['Choices'].forEach((curr)=>{otherCurrHtml +=`<option value="${curr}">${curr}</option>`})
  LocationDD.innerHTML = LocOptionHtml;
  StaffStatusDD.innerHTML = staffOptionHtml;
  othercurrDD.innerHTML = otherCurrHtml;
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
          Name: li.User.Title,
          FirstName: li.User.FirstName,
          LastName: li.User.LastName,
          Usermail: li.User.EMail,
          UserPersonalMail: li.PersonalEmail,
          JobTitle: li.User.JobTitle,
          Assistant: li.Assistant.Title,
          PhoneNumber: li.MobileNo != null ? li.MobileNo : "N/A",
          Location: li.SDGOffice,
          Title: li.stafffunction != null ? li.stafffunction : "N/A",
          Affiliation: li.SDGAffiliation,
          HAddLine: li.HomeAddLine,
          HAddCity: li.HomeAddCity,
          HAddState: li.HomeAddState,
          HAddPCode: li.HomeAddPCode,
          HAddPCountry: li.HomeAddCountry,
          ShortBio: li.ShortBio,
          Citizen: li.Citizenship,
          Industry: li.IndustryExp,
          Language: li.LanguageExp,
          SDGCourse: li.SDGCourses,
          Software: li.SoftwareExp,
          Membership: li.Membership,
          SpecialKnowledge: li.SpecialKnowledge,
          USDDaily: li.USDDailyRate,
          USDHourly: li.USDHourlyRate,
          EURDaily: li.EURDailyRate,
          EURHourly: li.EURHourlyRate,
          OtherCurr:li.OtherCurrency,
          OtherCurrDaily:li.ODailyRate,
          OtherCurrHourly:li.OHourlyRate,
          EffectiveDate: li.EffectiveDate,
          StaffStatus:li.StaffStatus,
          WorkSchedule:li.WorkingSchedule,
          ItemID:li.ID,
          LinkedInID:li.LinkedInLink,
          SignOther:li.signother,
          Child:li.children,
          HomeNo:li.HomeNo,
          EmergencyNo:li.EmergencyNo,
        });
      });
      getTableData();
      console.log(UserDetails);
      
    });
  UserProfileDetail();
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
          : "";
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
  let OfficeDetails = await sp.web
    .getList(listUrl + "SDGOfficeInfo")
    .items.get();
  // console.log(OfficeDetails);
  OfficeDetails.forEach((oDetail) => {
    OfficeTable += `<tr><td>${oDetail.Office}</td><td>${oDetail.Phone !="null"?oDetail.Phone.split(
      "^"
    ).join("</br>"):""}</td><td>${oDetail.Address != "null"?oDetail.Address.split("^").join(
      "</br>"
    ):""}</td></tr>`;
  });
  UserDetails.forEach((details) => {
    if (details.Affiliation == "Employee") {
      EmpTable += `<tr><td class="usernametag">${details.Name}</td><td>${details.FirstName}</td><td>${details.LastName}</td><td>${details.PhoneNumber}</td><td>${details.Location}</td><td>${details.JobTitle}</td><td>${details.Title}</td><td>${details.Assistant}</td></tr>`;
    }
    if (details.Affiliation == "Outside Consultant") {
      OutTable += `<tr><td class="usernametag">${details.Name}</td><td>${details.FirstName}</td><td>${details.LastName}</td><td>${details.PhoneNumber}</td><td>${details.Location}</td><td>${details.JobTitle}</td><td>${details.Title}</td><td>${details.Assistant}</td></tr>`;
    }
    if (details.Affiliation == "Affiliate") {
      AffTable += `<tr><td class="usernametag">${details.Name}</td><td>${details.FirstName}</td><td>${details.LastName}</td><td>${details.PhoneNumber}</td><td>${details.Location}</td><td>${details.JobTitle}</td><td>${details.Title}</td><td>${details.Assistant}</td></tr>`;
    }
    if (details.Affiliation == "Alumni") {
      AlumTable += `<tr><td class="usernametag">${details.Name}</td><td>${details.FirstName}</td><td>${details.LastName}</td><td>${details.PhoneNumber}</td><td>${details.Location}</td><td>${details.JobTitle}</td><td>${details.Title}</td><td>${details.Assistant}</td></tr>`;
    }
    AllDetailsTable += `<tr><td class="usernametag">${details.Name}</td><td>${details.FirstName}</td><td>${details.LastName}</td><td>${details.PhoneNumber}</td><td>${details.Location}</td><td>${details.JobTitle}</td><td>${details.Title}</td><td>${details.Assistant}</td></tr>`;
  });
  $("#SdhEmpTbody").html(EmpTable);
  $("#SdhOutsideTbody").html(OutTable);
  $("#SdhAffilateTbody").html(AffTable);
  $("#SdhAllumniTbody").html(AlumTable);
  $("#SdhAllPeopleTbody").html(AllDetailsTable);
  $("#SdgofficeinfoTbody").html(OfficeTable);
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
  UserProfileDetail();
}
const bindEmpTable = (options) => {
  $("#SdhEmpTable").DataTable(options);
};
const bindOutTable = (options) => {
  $("#SdhOutsideTable").DataTable(options);
};
const bindAffTable = (options) => {
  $("#SdhAffilateTable").DataTable(options);
};
const bindAlumTable = (options) => {
  $("#SdhAllumniTable").DataTable(options);
};
const bindAllDetailTable = (options) => {
  $("#SdhAllPeopleTable").DataTable(options);
};
const bindOfficeTable = (option) => {
  $("#SdgofficeinfoTable").DataTable(option);
};
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
  UserProfileDetail();
};

function startIt()
{
        var schema = {};
        schema['PrincipalAccountType'] = 'User,DL,SecGroup,SPGroup';
        schema['SearchPrincipalSource'] = 15;
        schema['ResolvePrincipalSource'] = 15;
        schema['AllowMultipleValues'] = true;
        schema['MaximumEntitySuggestions'] = 50;
        schema['Width'] = '280px';

        // Render and initialize the picker.
        // Pass the ID of the DOM element that contains the picker, an array of initial
        // PickerEntity objects to set the picker value, and a schema that defines
        // picker properties.
        SPClientPeoplePicker_InitStandaloneControlWrapper("peoplepickerText", null, schema);
}

const UserProfileDetail = async () => {
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

  
  //(<any>$("#peoplepickerText")).spPeoplePicker();

  const viewDir = document.querySelector(".view-directory");
  const editDir = document.querySelector(".edit-directory");
  const editbtn = document.querySelector(".edit-btn");
  const submitbtn = document.querySelector('#BtnSubmit');
  const cancelbtn = document.querySelector('#BtnCancel');
  let ItemID = 0;
  let OfficeAddArr = [];
  let office = await sp.web.getList(listUrl + "SDGOfficeInfo").items.get();
  office.forEach((off) => {
    OfficeAddArr.push({ OfficePlace: off.Office, OfficeFullAdd: off.Address });
  });
  let SelectedUser = "";
  const userpage = document.querySelector(".user-profile-page");
  const username = document.querySelectorAll(".usernametag");
  const sdhEmp = document.querySelector(".sdh-employee");
  
  const Edit = document.querySelector("#btnEdit");
const UserView = document.querySelector('.view-directory');
const UserEdit = document.querySelector('.edit-directory');
const StaffStatus = document.querySelector('#staffstatusDD');
const WorkscheduleSection = document.querySelector('#workscheduleSec');

  username.forEach( (btn) => {
    btn.addEventListener("click", async(e) => {
      if (!sdhEmp.classList.contains("hide")) {
        sdhEmp.classList.add("hide");
        userpage.classList.remove("hide");
      }
      SelectedUser = e.currentTarget["textContent"];
      let SelectedUserProfile = UserDetails.filter((li) => {
        return li.Name == SelectedUser;
      });
 
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
        let files = await sp.web.getFolderByServerRelativeUrl(`BiographyDocument/${SelectedUserProfile[0].Usermail}`).files.get();
        console.log(files);
        files.forEach((file)=>{
          console.log(file.Name.split('.').pop());
          
          if(file.Name.split('.').pop() == "doc" || file.Name.split('.').pop() =="docx"){
            filesHtml += `<div class="doc-section"><span class="word-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank">${file.Name}</a></div>`
          }else
          if(file.Name.split('.').pop() == "xlsx"){
            filesHtml += `<div class="doc-section"><span class="excel-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank">${file.Name}</a></div>`
          }else
          if(file.Name.split('.').pop() == "png" || file.Name.split('.').pop() =="jpg" ||file.Name.split('.').pop() == "jpeg"){
            filesHtml += `<div class="doc-section"><span class="pic-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank">${file.Name}</a></div>`
          }else{
            filesHtml += `<div class="doc-section"><span class="new-doc"></span><a href='${file.ServerRelativeUrl}' target="_blank">${file.Name}</a></div>`
          }
        })
        
         
  
      let billingRateHtml = "";
      if (SelectedUserProfile[0].USDDaily) {
        billingRateHtml += `<div class="billing-rates"><label>USD Daily Rates</label><div class="usd-daily-rate lblBlue" id="UsdDailyRate">${SelectedUserProfile[0].USDDaily}</div></div><div class="billing-rates"><label>USD Hourly Rates</label><div class="usd-hourly-rate lblBlue" id="UsdHourlyRate">${SelectedUserProfile[0].USDHourly}</div></div>`;
      }
      
      if (SelectedUserProfile[0].EURDaily != null) {
        billingRateHtml += `<div class="billing-rates"><label>EUR Daily Rates</label><div class="eur-daily-rate lblBlue" id="EURDailyRate">${SelectedUserProfile[0].EURDaily}</div></div><div class="billing-rates"><label>EUR Hourly Rates</label><div class="eur-hourly-rate lblBlue" id="EURHourlyRate">${SelectedUserProfile[0].EURHourly}</div></div>`;
      }
      
      if(SelectedUserProfile[0].OtherCurrDaily != null){
        billingRateHtml += `<div class="billing-rates"><label>${SelectedUserProfile[0].OtherCurr} Daily Rates</label><div class="eur-daily-rate lblBlue" id="oDailyRate">${SelectedUserProfile[0].OtherCurrDaily}</div></div><div class="billing-rates"><label>${SelectedUserProfile[0].OtherCurr} Hourly Rates</label><div class="eur-hourly-rate lblBlue" id="oHourlyRate">${SelectedUserProfile[0].OtherCurrHourly}</div></div>`
      }
      if (SelectedUserProfile[0].EffectiveDate != null) {
        billingRateHtml += ` <div class="billing-effective-date"><label>Effective Date</label><div class="effective-date lblBlue" id="EffectiveDate">${new Date(
          SelectedUserProfile[0].EffectiveDate
        ).toLocaleDateString()}</div></div>`;
      } 
      ItemID = SelectedUserProfile[0].ItemID; 
      
      
      $("#user-Designation").html(SelectedUserProfile[0].Affiliation);
      $("#user-job-title").html(SelectedUserProfile[0].JobTitle);
      $("#user-location").html(SelectedUserProfile[0].Location);
      $("#user-office").html(SelectedUserProfile[0].Location);
      $("#user-phone").html(SelectedUserProfile[0].PhoneNumber);
      $("#user-mail").html(SelectedUserProfile[0].Usermail);
      $('#personal-mail').html(SelectedUserProfile[0].UserPersonalMail);
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
      $("#shortbio").html(SelectedUserProfile[0].ShortBio);
      $("#citizenship").html(SelectedUserProfile[0].Citizen);
      $("#IndustryExp").html(SelectedUserProfile[0].Industry);
      $("#LanguageExp").html(SelectedUserProfile[0].Language);
      $("#SDGCourse").html(SelectedUserProfile[0].SDGCourse);
      $("#SoftwareExp").html(SelectedUserProfile[0].Software);
      $("#MembershipExp").html(SelectedUserProfile[0].Membership);
      $("#SpecialKnowledge").html(SelectedUserProfile[0].SpecialKnowledge);
      $("#BillingRateDetails").html(billingRateHtml);
      $('#bioAttachment').html(filesHtml);
      $('#staffStatus').html(SelectedUserProfile[0].StaffStatus);
      $('#workscheduleViewSec').html(SelectedUserProfile[0].StaffStatus == "Part-time"?`<h5>Work Schedule</h5>
      <p class="lblRight" id="workSchedule">${SelectedUserProfile[0].WorkSchedule}</p>`:"");

      // Edit Section
      Edit.addEventListener('click',(btn)=>{
        if(!UserView.classList.contains("hide")){
          UserView.classList.add('hide')
          UserEdit.classList.remove('hide');
          Edit.classList.add('hide');
        }else{
          UserEdit.classList.remove('hide');
          Edit.classList.add('hide');
        }
        $('#PAddLineE').val(SelectedUserProfile[0].HAddLine);
        $('#PAddCityE').val(SelectedUserProfile[0].HAddCity);
        $('#PAddStateE').val(SelectedUserProfile[0].HAddState);
        $('#PAddPCodeE').val(SelectedUserProfile[0].HAddPCode);
        $('#PAddCountryE').val(SelectedUserProfile[0].HAddPCountry);
        $('#Eshortbio').val(SelectedUserProfile[0].ShortBio);
        $('#EIndustry').val(SelectedUserProfile[0].Industry);
        $('#ELanguage').val(SelectedUserProfile[0].Language);
        $('#ESDGCourse').val(SelectedUserProfile[0].SDGCourse);
        $('#ESoftwarExp').val(SelectedUserProfile[0].Software);
        $('#EMembership').val(SelectedUserProfile[0].Membership);
        $('#ESKnowledge').val(SelectedUserProfile[0].SpecialKnowledge);
        $('#citizenshipE').val(SelectedUserProfile[0].Citizen);
        $("#linkedInID").val(SelectedUserProfile[0].LinkedInID);
        $('#mobileno').val(SelectedUserProfile[0].PhoneNumber);
        $('#children').val(SelectedUserProfile[0].Child);
        $('#significantOther').val(SelectedUserProfile[0].SignOther);
        $('#USDDailyEdit').val(SelectedUserProfile[0].USDDaily);
        $('#USDHourlyEdit').val(SelectedUserProfile[0].USDHourly);
        $('#EURDailyEdit').val(SelectedUserProfile[0].EURDaily);
        $('#EURHourlyEdit').val(SelectedUserProfile[0].EURHourly);
        $("#assisstantName").val(SelectedUserProfile[0].Assistant);
        $("#personalmailID").val(SelectedUserProfile[0].UserPersonalMail);
        $('#homeno').val(SelectedUserProfile[0].HomeNo); 
        $('#emergencyno').val(SelectedUserProfile[0].EmergencyNo);
        $('#workLocationDD').val(SelectedUserProfile[0].Location);
        $('#staffstatusDD').val(SelectedUserProfile[0].StaffStatus);
        $('#othercurrDD').val(SelectedUserProfile[0].OtherCurr);
        $('#ODailyEdit').val(SelectedUserProfile[0].OtherCurrDaily);
        $('#OHourlyEdit').val(SelectedUserProfile[0].OtherCurrHourly);
        var  finalmonth:any="";
        var dd=new Date(SelectedUserProfile[0].EffectiveDate).getDate();
        var mm=new Date(SelectedUserProfile[0].EffectiveDate).getMonth()+1;
        mm<10?finalmonth="0"+mm:finalmonth=mm
        var yyyy=new Date(SelectedUserProfile[0].EffectiveDate).getFullYear()
var  dateformat=yyyy+"-"+finalmonth+"-"+dd;
$('#EffectiveDateEdit').val(dateformat);
if($("#staffstatusDD").val() == "Part-time"){
  $("#workscheduleSec").classList.remove("hide");
}else{
  $("#workscheduleSec").val("")
}
      });
    });
  });

  $('#USDDailyEdit').change(()=>{
    var usdvalue:any =$('#USDDailyEdit').val();
    var finalusdval=usdvalue/8;
    $('#USDHourlyEdit').val(finalusdval);
  });
  $('#EURDailyEdit').change(()=>{
    var eurdaily:any = $('#EURDailyEdit').val();
    var finaleurval = eurdaily/8;
    $('#EURHourlyEdit').val(finaleurval)
  });
  $('#ODailyEdit').change(()=>{
    var ovalue:any = $('#ODailyEdit').val();
    var finalovalue = ovalue/8;
    $('#OHourlyEdit').val(finalovalue)
  })
  submitbtn.addEventListener('click',async()=>{
    var dispTitle = "APickerField";      
    var pickerDiv = $("[id$='peoplepickerText'][title='" + dispTitle + "']");      
    var peoplePicker = SPClientPeoplePicker.SPClientPeoplePickerDict;
    var userInfo = peoplePicker.peoplepickerText_TopSpan.GetAllUserInfo()
    const loginName = userInfo[0].EntityData.Email;
    const profile =  await sp.web.siteUsers.getByEmail(loginName).get();

    console.log(profile.Id);
    console.log(userInfo);
    
const update = await sp.web.getList(listUrl + "StaffDirectory").items.getById(ItemID).update({
  Title: "SDG User Info",
  PersonalEmail:$("#personalmailID").val(),
  MobileNo:$('#mobileno').val(),
  HomeNo:$('#homeno').val(),
  EmergencyNo:$('#emergencyno').val(),
  HomeAddLine:$('#PAddLineE').val(),
  HomeAddCity:$('#PAddCityE').val(),
  HomeAddState:$('#PAddStateE').val(),
  HomeAddPCode:$('#PAddPCodeE').val(),
  HomeAddCountry:$('#PAddCountryE').val(),
  IndustryExp:$('#EIndustry').val(),
  LanguageExp:$('#ELanguage').val(),
  SDGCourses:$('#ESDGCourse').val(),
  SoftwareExp:$('#ESoftwarExp').val(),
  Membership:$('#EMembership').val(),
  SpecialKnowledge:$('#ESKnowledge').val(),
  Citizenship:$('#citizenshipE').val(),
  ShortBio:$('#Eshortbio').val(),
  USDDailyRate:$('#USDDailyEdit').val(),
  USDHourlyRate:$('#USDHourlyEdit').val(),
  EURDailyRate:$('#EURDailyEdit').val(),
  EURHourlyRate:$('#EURHourlyEdit').val(),
  OtherCurrency:$('#othercurrDD').val(),
  ODailyRate:$('#ODailyEdit').val(),
  OHourlyRate:$('#OHourlyEdit').val(),
  EffectiveDate:$('#EffectiveDateEdit').val(),
  signother:$('#significantOther').val(),
  children:$('#children').val(),
  WorkingSchedule:$('#workSchedule').val(),
  SDGOffice:$('#workLocationDD').val(),
  StaffStatus:$('#staffstatusDD').val(),
  AssistantId:profile.Id
});
// location.reload();
  })
  cancelbtn.addEventListener('click',()=>{
    viewDir.classList.remove("hide");
    editDir.classList.add("hide");
    editbtn.classList.remove("hide");
  })
  
};
