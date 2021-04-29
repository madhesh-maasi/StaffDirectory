import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import styles from "./StaffdirectoryWebPart.module.scss";
import * as strings from "StaffdirectoryWebPartStrings";
import { SPComponentLoader } from "@microsoft/sp-loader";
import "../../ExternalRef/CSS/style.css";
import * as $ from "jquery";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import "@pnp/graph/photos";
import "@pnp/sp/profiles";
SPComponentLoader.loadScript(
  "https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.4.min.js"
  // "https://code.jquery.com/jquery-3.5.1.js"
);
SPComponentLoader.loadCss(
  "https://cdn.jsdelivr.net/npm/select2@4.1.0-beta.1/dist/css/select2.min.css"
);
SPComponentLoader.loadCss(
  "https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"
);
import "datatables";
import { _SharePointQueryableInstance } from "@pnp/sp/sharepointqueryable";
require("datatables.net-dt");
require("datatables.net-rowgroup-dt");
// import "../../../node_modules/datatables/media/css/jquery.dataTables.min.css";
SPComponentLoader.loadCss(
  "https://cdn.datatables.net/1.10.24/css/jquery.dataTables.css"
);

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
var table;
let UserDetails = [];
var oDataTable: any;

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
           <div  data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseThree" aria-expanded="false" aria-controls="collapseThree"><span class="nav-icon affli"></span>Affilates</div>
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
         <div data-toggle="collapse" class="clsToggleCollapse" data-target="#collapseFour" aria-expanded="false" aria-controls="collapseFour"><span class="nav-icon sdh-alumini"></span>SDG Allumini</div>
       </div>
       <div id="collapseFour" class="clsCollapse collapse" aria-labelledby="headingFour" data-parent="#accordionExample">
         <div class="card-body">
         <div class="filter-section">
         <ul>
         <li><a href="#" class="SDHAlumniLastName">By Last Name</a></li>
         <li><a href="#" class="SDHAlumniFirstName">By First Name</a></li>
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
       <div class="card-header nav-items" id="headingSix">
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
     <div class="title-font" id="user-mail">madheshmaasi3@gmail.com</div>
     </div>
     </div>
     </div>
     </div>
     </div>
     </div>    
     </div> 
     </div>`;

    const username = document.querySelectorAll(".usernametag");
    const userpage = document.querySelector(".user-profile-page");
    const tableSection = document.querySelector(".sdh-employee");
    $(".clsToggleCollapse").click(function () {
      $(".clsCollapse").each(function () {
        $(this).removeClass("in").attr("style", "");
      });
      $(this).next("div").addClass("in");
    });
    onLoadData();
    ActiveSwitch();
    $(".SDHEmployee").click(() => {
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindEmpTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindEmpTable(options);
      }
    });
    $(".OutsidConsultant").click(() => {
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindOutTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindOutTable(options);
      }
    });
    $(".SDHAffiliates").click(() => {
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindAffTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindAffTable(options);
      }
    });
    $(".SDHAlumini").click(() => {
      if (tableSection.classList.contains("hide")) {
        tableSection.classList.remove("hide");
        userpage.classList.add("hide");
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindAlumTable(options);
      } else {
        var options = {
          destroy: true,
          order: [[0, "asc"]],
        };
        bindAlumTable(options);
      }
    });
    $(".SDHShowAll").click(() => {
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
  sp.web.lists
    .getByTitle("StaffDirectory")
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
      // console.log(listitem);
      listitem.forEach((li) => {
        UserDetails.push({
          Name: li.User.Title,
          FirstName: li.User.FirstName,
          LastName: li.User.LastName,
          Usermail: li.User.EMail,
          JobTitle: li.User.JobTitle,
          Assistant: li.Assistant.Title,
          PhoneNumber: li.MobileNo != null ? li.MobileNo : "N/A",
          Location: li.SDGOffice,
          Title: li.stafffunction != null ? li.stafffunction : "N/A",
          Affiliation: li.SDGAffiliation,
        });
      });
      getTableData();
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
          : "";
      });
    });
  });
};

function getTableData() {
  let EmpTable = "";
  let OutTable = "";
  let AffTable = "";
  let AlumTable = "";
  let AllDetailsTable = "";
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
  var options = {
    order: [[0, "asc"]],
  };
  bindEmpTable(options);
  bindOutTable(options);
  bindAffTable(options);
  bindAlumTable(options);
  bindAllDetailTable(options);
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

const UserProfileDetail = () => {
  let SelectedUser = "";
  const userpage = document.querySelector(".user-profile-page");
  const username = document.querySelectorAll(".usernametag");
  const sdhEmp = document.querySelector(".sdh-employee");
  username.forEach((btn) => {
    btn.addEventListener("click", (e) => {
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
      $("#user-Designation").html(SelectedUserProfile[0].Affiliation);
      $("#user-job-title").html(SelectedUserProfile[0].JobTitle);
      $("#user-location").html(SelectedUserProfile[0].Location);
      $("#user-office").html(SelectedUserProfile[0].Location);
      $("#user-phone").html(SelectedUserProfile[0].PhoneNumber);
      $("#user-mail").html(SelectedUserProfile[0].Usermail);
      $("#UserProfileName").html(SelectedUserProfile[0].Name);
      $("#UserProfileEmail").html(
        `<span class="user-mail-icon"></span>${SelectedUserProfile[0].Usermail}`
      );
    });
  });
};
