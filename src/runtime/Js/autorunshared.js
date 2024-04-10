// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on Outlook on web, on Windows, and on Mac (new UI preview).

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */
function checkSignature(eventObj) {
  console.log("Office.context", Office.context)
  
  getCurrentUser(eventObj);
}

function getCurrentUser(eventObj){
  var userDetails = {
    "UserDetails": [
      {
        "Department": "Contracts & Procurement",
        "FirstName": "Albina",
        "LastName": "Bezdetko",
        "MobilePhone": "+61 459 330 065",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 459 330 065",
        "Title": "Associate - Contracts & Procurement",
        "Email": "albina.bezdetko@lidiargroup.com.au"
      },
      {
        "Department": "Northern Territory",
        "FirstName": "Arthur",
        "LastName": "Dimitropoulos",
        "MobilePhone": "+61 408 508 677",
        "Office": "Level 1 - Suite 52, 48 - 50 Smith Street, Darwin, NT, 0800.",
        "PhoneNumber": "+61 408 508 677",
        "Title": "Senior Associate",
        "Email": "arthur.dimitropoulos@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Behshad",
        "LastName": "Bordbar",
        "MobilePhone": "+61 434 137 985",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 434 137 985",
        "Title": "Associate - Project Controls",
        "Email": "behshad.bordbar@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Blair",
        "LastName": "Barton",
        "MobilePhone": "+61 400 612 231",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 400 612 231",
        "Title": "Associate - Senior Project Manager",
        "Email": "blair.barton@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Candy",
        "LastName": "Lopez",
        "MobilePhone": "+61 432 251 184",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 432 251 184",
        "Title": "Associate - Project Engineer",
        "Email": "candy.lopez@lidiargroup.com.au"
      },
      {
        "Department": "Contracts & Procurement",
        "FirstName": "Clara",
        "LastName": "Nyamandi",
        "MobilePhone": "+61 434 635 462",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 434 635 462",
        "Title": "Associate - Contracts & Procurement",
        "Email": "clara.nyamandi@lidiargroup.com.au"
      },
      {
        "Department": "Corporate",
        "FirstName": "Clarisse",
        "LastName": "Baldoza",
        "MobilePhone": "+63 927 350 0861",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+63 927 350 0861",
        "Title": "Associate - Accounts and Finance",
        "Email": "clarisse.baldoza@lidiargroup.com.au"
      },
      {
        "Department": "d",
        "FirstName": "Conrad",
        "LastName": "Hine",
        "MobilePhone": "+61 421 502 782",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 421 502 782",
        "Title": "Associate - Business Analyst",
        "Email": "conrad.hine@lidiargroup.com.au"
      },
      {
        "Department": "Health, Safety & Environment",
        "FirstName": "Corey",
        "LastName": "Kolar",
        "MobilePhone": "+61 403 256 313",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 403 256 313",
        "Title": "Associate - Health & Safety Consultant",
        "Email": "corey.kolar@lidiargroup.com.au"
      },
      {
        "FirstName": "Daniel",
        "LastName": "Stone",
        "MobilePhone": "+61 468 780 297",
        "Office": "Level 3, 240 Queen Street",
        "PhoneNumber": "+61 468 780 297",
        "Title": "Associate - Health & Safety Advisor",
        "Email": "daniel.stone@lidiargroup.com.au"
      },
      {
        "FirstName": "Darren",
        "LastName": "Cave",
        "MobilePhone": "+61 418 619 137",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 418 619 137",
        "Title": "Partner",
        "Email": "darren.cave@lidiargroup.com.au"
      },
      {
        "Department": "Lidiar Advisory",
        "FirstName": "David",
        "LastName": "Plowman",
        "MobilePhone": "+61 452 000 450",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 452 000 450",
        "Title": "Consulting Executive",
        "Email": "david.plowman@lidiargroup.com.au"
      },
      {
        "Department": "Corporate",
        "FirstName": "Erwin",
        "LastName": "Domantay",
        "MobilePhone": "+63 9985869838",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+63 9985869838",
        "Title": "Associate - SharePoint Developer",
        "Email": "erwin.domantay@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Frederico",
        "LastName": "de Souza",
        "MobilePhone": "+61 447 307 098 ",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 447 307 098 ",
        "Title": "Associate - Project Engineer",
        "Email": "frederico.desouza@lidiargroup.com.au"
      },
      {
        "FirstName": "Geoff",
        "LastName": "Saunders",
        "MobilePhone": "+61 448 735 253",
        "Office": "Brisbane",
        "Title": "Associate - Project Engineer",
        "Email": "geoff.saunders@lidiargroup.com.au"
      },
      {
        "Department": "Northern Territory",
        "FirstName": "Georgie",
        "LastName": "Coles",
        "MobilePhone": "+61 466 012 587",
        "Office": "Level 1 - Suite 52, 48 - 50 Smith Street, Darwin, NT, 0800.",
        "PhoneNumber": "+61 466 012 587",
        "Title": "Associate - Environmental Consultant",
        "Email": "georgie.coles@lidiargroup.com.au"
      },
      {
        "FirstName": "Georgina",
        "LastName": "Poole",
        "MobilePhone": "+61 437 234 059 ",
        "Office": "Brisbane",
        "Email": "georgina.poole@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Gerald",
        "LastName": "Guillet",
        "MobilePhone": "+61 499 453 535",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 499 453 535",
        "Title": "Associate - Project Manager",
        "Email": "gerald.guillet@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Ishara",
        "LastName": "Kadugalla",
        "MobilePhone": "+61 403 481 211",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 403 481 211",
        "Title": "Associate - Quantity Surveyor",
        "Email": "ishara.kadugalla@lidiargroup.com.au"
      },
      {
        "Department": "Corporate",
        "FirstName": "Jessica",
        "LastName": "Romero",
        "MobilePhone": "+63 965 644 7570",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+63 965 644 7570",
        "Title": "Associate - Timesheet Coordinator",
        "Email": "jessica.romero@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Jessica",
        "LastName": "Talik",
        "MobilePhone": "+61 0450 646 118",
        "Office": "Level 3, 240 Queen Street",
        "PhoneNumber": "+61 0450 646 118",
        "Title": "Associate - Quantity Surveyor",
        "Email": "jessica.talik@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Joseph",
        "LastName": "Boylan",
        "MobilePhone": "+61 419 846 746",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 419 846 746",
        "Title": "Associate - Project Manager",
        "Email": "joseph.boylan@lidiargroup.com.au"
      },
      {
        "Department": "Health & Safety",
        "FirstName": "Julio",
        "LastName": "Bará",
        "MobilePhone": "+61 422 229 710",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 422 229 710",
        "Title": "Senior Associate - Health & Safety",
        "Email": "julio.bara@lidiargroup.com.au"
      },
      {
        "Department": "Contracts & Procurement",
        "FirstName": "Justin",
        "LastName": "Bone",
        "MobilePhone": "+61 473 015 451",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 473 015 451",
        "Title": "Senior Associate - Contracts & Procurement",
        "Email": "justin.bone@lidiargroup.com.au"
      },
      {
        "Department": "Health & Safety",
        "FirstName": "Kieren",
        "LastName": "Thomas",
        "MobilePhone": "+61 413 466 373",
        "Office": "Level 3, 240 Queen Street",
        "PhoneNumber": "+61 413 466 373",
        "Title": "Associate - Health & Safety",
        "Email": "kieren.thomas@lidiargroup.com.au"
      },
      {
        "Department": "Management",
        "FirstName": "Lachlan",
        "LastName": "Winterbotham",
        "MobilePhone": "+61 437 234 059",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 437 234 059",
        "Title": "Partner",
        "Email": "lachlan.winterbotham@lidiargroup.com.au"
      },
      {
        "Department": "Corporate",
        "FirstName": "Lara",
        "LastName": "Lindsay",
        "MobilePhone": "+61 402 789 654",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 402 789 654",
        "Title": "Associate - Business Operations",
        "Email": "lara.lindsay@lidiargroup.com.au"
      },
      {
        "Department": "Contracts & Procurement",
        "FirstName": "Laura",
        "LastName": "Vacca",
        "MobilePhone": "+61 435 031 664",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 435 031 664",
        "Title": "Associate - Contracts & Procurement",
        "Email": "laura.vacca@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Lily",
        "LastName": "Snezhina",
        "MobilePhone": "+61 481 123 773",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 481 123 773",
        "Title": "Associate - Business Analyst",
        "Email": "lily.snezhina@lidiargroup.com.au"
      },
      {
        "Department": "Northern Territory",
        "FirstName": "Louis",
        "LastName": "Alderdice",
        "MobilePhone": "+61 416 050 521",
        "Office": "Level 1 - Suite 52, 48 - 50 Smith Street, Darwin, NT, 0800.",
        "PhoneNumber": "+61 416 050 521",
        "Title": "Associate - Project Engineer",
        "Email": "louis.alderdice@lidiargroup.com.au"
      },
      {
        "Department": "d",
        "FirstName": "Marianys",
        "LastName": "Diaz",
        "MobilePhone": "+61 410 682 547",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 410 682 547",
        "Title": "Associate - Project Engineer",
        "Email": "marianys.diaz@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Michael",
        "LastName": "Beaven",
        "MobilePhone": "+61 413 880 198",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 413 880 198",
        "Title": "Associate - Project Controls",
        "Email": "michael.beaven@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Natalie",
        "LastName": "Robinson",
        "MobilePhone": "+61 424 274 446",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 424 274 446",
        "Title": "Associate - Environmental Consultant",
        "Email": "natalie.robinson@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Natig",
        "LastName": "Nabiyev",
        "MobilePhone": "+61 490 409 520",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 490 409 520",
        "Title": "Project Engineer",
        "Email": "natig.nabiyev@lidiargroup.com.au"
      },
      {
        "Department": "Management",
        "FirstName": "Niall",
        "LastName": "Callan",
        "MobilePhone": "+61 405 113 793",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 405 113 793",
        "Title": "Partner",
        "Email": "niall.callan@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Olivia",
        "LastName": "Hamilton",
        "MobilePhone": "+61 450 612 069",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 450 612 069",
        "Title": "Associate - Social Performance Analyst",
        "Email": "olivia.hamilton@lidiargroup.com.au"
      },
      {
        "Department": "Advisory",
        "FirstName": "Pierre",
        "LastName": "Vermeulen",
        "MobilePhone": "+61 451 208 297 ",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 451 208 297 ",
        "Title": "Associate - Commercial",
        "Email": "pierre.vermeulen@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Rebecca",
        "LastName": "Patrick",
        "MobilePhone": "+61 408 064 089",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 408 064 089",
        "Title": "Associate - Creative & Design",
        "Email": "rebecca.patrick@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Rick",
        "LastName": "Winsor",
        "MobilePhone": "+61 400 678 844",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 400 678 844",
        "Title": "Associate - Project Engineer",
        "Email": "rick.winsor@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Roshan",
        "LastName": "Mathew",
        "MobilePhone": "+61 424 134 075",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 424 134 075",
        "Title": "Associate - Project Engineer",
        "Email": "roshan.mathew@lidiargroup.com.au"
      },
      {
        "Department": "Health & Safety",
        "FirstName": "Savio",
        "LastName": "Pereira",
        "MobilePhone": "+61 450 670 331",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 450 670 331",
        "Title": "Associate - Health & Safety",
        "Email": "savio.pereira@lidiargroup.com.au"
      },
      {
        "FirstName": "Scott",
        "LastName": "John",
        "MobilePhone": "+61 409 130 453",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "Title": "Consultant - Reciprocity - Organisational Performance Specialist",
        "Email": "scott.john@lidiargroup.com.au"
      },
      {
        "Department": "Contracts & Procurement",
        "FirstName": "Sean",
        "LastName": "Casey",
        "MobilePhone": "+61 428 916 151",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 428 916 151",
        "Title": "Associate - Contracts & Procurement",
        "Email": "sean.casey@lidiargroup.com.au"
      },
      {
        "Department": "Management",
        "FirstName": "Shane",
        "LastName": "Synnott",
        "MobilePhone": "+61 404 812 454",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 404 812 454",
        "Title": "Advisory Board Member",
        "Email": "shane.synnott@lidiargroup.com.au"
      },
      {
        "Department": "Environmental",
        "FirstName": "Steve",
        "LastName": "Onogbo",
        "MobilePhone": "+61 410 826 339",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 410 826 339",
        "Title": "Associate - Environment Consultant",
        "Email": "steve.onogbo@lidiargroup.com.au"
      },
      {
        "Department": "Project Management Services",
        "FirstName": "Yoshi",
        "LastName": "Lim",
        "MobilePhone": "+61 492 455 580",
        "Office": "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.",
        "PhoneNumber": "+61 492 455 580",
        "Title": "Associate - Project Engineer",
        "Email": "yoshi.lim@lidiargroup.com.au"
      }
    ]
  };


  // https://lidiargroup.sharepoint.com/sites/IntegratedManagementSystem/SiteAssets/UserDetails.json
  // $.getJSON("https://raw.githubusercontent.com/ejdomantay/lidiar-group-signature/main/src/runtime/UserDetails.json", function(response) {
      
  // });


    
  const curUser = userDetails.filter(x => x.Email == Office.context.mailbox.userProfile.emailAddress)[0];
    setSignatureTemplate({
      Title: curUser.Title,
      Phones: [{ Type: "Business", Number: curUser.PhoneNumber}],
      OfficeLocation: curUser.Office
      }, eventObj);

  // Office.context.mailbox.getCallbackTokenAsync({
  //   isRest: true
  // }, function (result) {
  //   console.log(result);  

  //   //const apiUrl = Office.context.mailbox.restUrl + "/v2.0/Users('" + Office.context.mailbox.userProfile.emailAddress + "')";
  //   //const apiUrl = Office.context.mailbox.restUrl + "/v1.0/Users('" + Office.context.mailbox.userProfile.emailAddress + "')/Contacts";
  //   // const apiUrl = Office.context.mailbox.restUrl + "/beta/Users('" + Office.context.mailbox.userProfile.emailAddress + "')/people?$top=200";
  //   const apiUrl = Office.context.mailbox.restUrl + "/v2.0/me/people?$top=200"
  //   $.ajax({  
  //     method: 'GET',  
  //     url: apiUrl,  
  //     headers: {  
  //         'Authorization': 'Bearer ' + result.value,  
  //         'Content-Type': 'application/json'  
  //     },  
  //   }).success(function(response) {   
  //       const curUser = response.value.filter(x => x.UserPrincipalName == Office.context.mailbox.userProfile.emailAddress)[0];
  //       console.log("response.value", curUser);
  //       console.log("curUser", curUser);
  //       setSignatureTemplate(curUser, eventObj);
  //   }).error(function(error) {
  //      console.log(error);
  //      setSignatureTemplate({
  //       Title: "",
  //       Phones: [{ Type: "Business", Number: ""}],
  //       OfficeLocation: ""
  //      }, eventObj);
  //   });

    
  // });
}


function setSignatureTemplate(curUser, eventObj){
  
  const busNo = curUser.Phones.filter(x => x.Type == "Business")[0];
  const emailTemplate = 
  
  '<span style="font-size:14px"><b>'+ Office.context.mailbox.userProfile.displayName +'</b></span>'+
  '<br />'+
  '<span style="font-size:14px">'+ curUser.Title +'<span>'+
  '<br />'+
  '<br />'+
  '<table style="border:0;border-spacing:0;" cellspacing="0">'+
    '<tr>'+
      '<td style="padding-right: 20px;">'+
        '<img  height="82" width="120" src="https://raw.githubusercontent.com/ejdomantay/lidiar-group/main/Lidiar%20Main%20Logo.png"></img>'+
      '</td>'+
      '<td>'+
        '<table style="border:0;border-spacing:0;" cellspacing="0">'+		
          '<tr>'+
            '<td style="background-color: red; width: 5px; height: 95px;">'+
          '</td>'+
          '</tr>'+
        '</table>'+
      '</td>'+
      '<td style="padding-left:5px;">'+
        '<table style="border:0;border-spacing:0; font-size:14px; line-height: 16px;" cellspacing="0">	'+	
          '<tr>'+
            '<td>'+
                '<span><b>Lidiar Group Pty Ltd</b></span>'+
             '</td>'+
          '</tr>'+
          '<tr>'+
            '<td>'+
                '<span style="color:red">m.</span>'+
                '<span> '+ (busNo ? busNo.Number : "") +'</span>'+
                '<span> | <span>'+
                '<span style="color:red">e.</span>'+
                '<span> '+ Office.context.mailbox.userProfile.emailAddress +'</span>'+
            '</td>'+
          '</tr>'+
          '<tr>'+
            '<td>'+
              '<span style="color:red">o.</span>'+
              '<span> '+ (curUser.OfficeLocation ? curUser.OfficeLocation : "Level 5, 144 Edward Street, Brisbane, Queensland, 4000.") +'</span>'+
                
            '</td>'+
          '</tr>'+
          '<tr>'+
          '<td>'+
            '<span style="color:red">w.</span> <span>www.lidiargroup.com.au</span>'+
                '</td>'+
            '</tr>'+
          '<tr>'+
            
          '<td>'+
            '<img style="" height="15" width="15" src="https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Microsoft_Office_Teams_%282018%E2%80%93present%29.svg/2203px-Microsoft_Office_Teams_%282018%E2%80%93present%29.svg.png"></img>'+
              
              '	<a href="https://teams.microsoft.com/l/chat/0/0?users='+ Office.context.mailbox.userProfile.emailAddress +'">'+
        
              ' Chat with me on Teams'+
                '	</a>'+
              '	</td>'+
            '</tr>'+	
          '</table>'+
        '</td>'+
      '</tr>'+
    '</table>'+
  '<div>'+
  '<p style="color:gray; font-size:13px;padding-top: 16px;">'+
        'Disclaimer: The information contained in this email is intended only for the use of the person(s) to whom it is addressed and may be confidential or contain privileged information. All information contained in this electronic communication is solely for the use of the individual(s) or entity to which it was addressed. If you are not the intended recipient you are hereby notified that any perusal, use, distribution, copying or disclosure is strictly prohibited. If you have received this email in error please immediately advise us by return email and delete the email without making a copy.'+
        '</p>'+
        '</div>'+
  '<div>'+
  '<p style="font-size:13px;padding-top: 10px;">'+
        'Please consider the environment before printing this email.'+
        '</p>'+
    '</div>'
 


    if(Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Appointment)
    {
      Office.context.mailbox.item.body.setAsync(
        "<br/><br/>" + emailTemplate,
        {
          coercionType: "html",
          asyncContext: eventObj,
        },
        function (asyncResult) {
    
          asyncResult.asyncContext.completed();
        }
      );
    }
    else{
    Office.context.mailbox.item.body.setSignatureAsync(
      emailTemplate,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

function requestToken() {  
  console.log(Office.context); 
  Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
        const token = result.value;
        console.log("token", token); 
    } else {
        console.log("Error obtaining token", result.error);
    }
});
}

function getUserInformation(token, mailAddress) {  
  $.ajax({  
      method: 'GET',  
      url: "https://graph.microsoft.com/v1.0/users/" + mailAddress,  
      headers: {  
          'Authorization': 'Bearer ' + token,  
          'Content-Type': 'application/json'  
      },  
  }).success(function(response) {  
      console.log(response);  
  }).error(function(error) {});  
} 
/**
 * For Outlook on Windows and on Mac only. Insert signature into appointment or message.
 * Outlook on Windows and on Mac can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
  let template_name = get_template_name(compose_type);
  let signature_info = get_signature_info(template_name, user_info);
  addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
  if (is_valid_data(signatureDetails.logoBase64) === true) {
    //If a base64 image was passed we need to attach it.
    Office.context.mailbox.item.addFileAttachmentFromBase64Async(
      signatureDetails.logoBase64,
      signatureDetails.logoFileName,
      {
        isInline: true,
      },
      function (result) {
        //After image is attached, insert the signature
        Office.context.mailbox.item.body.setSignatureAsync(
          signatureDetails.signature,
          {
            coercionType: "html",
            asyncContext: eventObj,
          },
          function (asyncResult) {
            asyncResult.asyncContext.completed();
          }
        );
      }
    );
  } else {
    //Image is not embedded, or is referenced from template HTML
    Office.context.mailbox.item.body.setSignatureAsync(
      signatureDetails.signature,
      {
        coercionType: "html",
        asyncContext: eventObj,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
  Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
    type: "insightMessage",
    message: "Please set your signature with the Office Add-ins sample.",
    icon: "Icon.16x16",
    actions: [
      {
        actionType: "showTaskPane",
        actionText: "Set signatures",
        commandId: get_command_id(),
        contextData: "{''}",
      },
    ],
  });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
  if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
  if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
  return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
  if (template_name === "templateB") return get_template_B_info(user_info);
  if (template_name === "templateC") return get_template_C_info(user_info);
  return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
  if (Office.context.mailbox.item.itemType == "appointment") {
    return "MRCS_TpBtn1";
  }
  return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */
function get_template_A_info(user_info) {
  const logoFileName = "sample-logo.png";
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Embed the logo using <img src='cid:...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='cid:" +
    logoFileName +
    "' alt='MS Logo' width='24' height='24' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += is_valid_data(user_info.job) ? user_info.job + "<br/>" : "";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  // return object with signature HTML, logo image base64 string, and filename to reference it with.
  return {
    signature: str,
    logoBase64:
      "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC",
    logoFileName: logoFileName,
  };
}

/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += "<table>";
  str += "<tr>";
  // Reference the logo using a URI to the web server <img src='https://...
  str +=
    "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
  str += "<td style='padding-left: 5px;'>";
  str += "<strong>" + user_info.name + "</strong>";
  str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
  str += "<br/>";
  str += user_info.email + "<br/>";
  str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
  str += "</td>";
  str += "</tr>";
  str += "</table>";

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
  let str = "";
  if (is_valid_data(user_info.greeting)) {
    str += user_info.greeting + "<br/>";
  }

  str += user_info.name;

  return {
    signature: str,
    logoBase64: null,
    logoFileName: null,
  };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
  return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);
