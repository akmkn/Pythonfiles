var lsincrementor;
lsincrementor = 0;
ApvlTypeincrementor = 1;


//Associating the script name of approval type to the dynamic added rows
function ApvlTypeScriptListLoad() {
   	var ApvlTypeData =(Scriptlist);
	ApvlTypeItem = "";
	for (item in ApvlTypeData)
	{
		var  scriptname=ApvlTypeData[item];
		ApvlTypeItem  = ApvlTypeItem + "<option value="+scriptname+">"+scriptname+"</option>";
	}
	return ApvlTypeItem;
}

//hide and  display the RM sanity table
function showhide()
{
    var checkbox = document.getElementById("RMSanity");
    var hiddeninputs = document.getElementsByClassName("hideRMScriptTable");
    for (var i=0;i!=hiddeninputs.length;i++)
        {
        if (checkbox.checked)
            {
            hiddeninputs[i].style.display = "table"
            }
        else
            {
            hiddeninputs[i].style.display = "none"
            }

        }
}

//hide and display Approval Type table
function showhideapprovaltypetable()
{
    var approvaltypechkbox = document.getElementById("approvaltype");
    var aprvltypehiddeninputs = document.getElementsByClassName("hideApprovalTypeTable");
    for (var i=0;i!=aprvltypehiddeninputs.length;i++)
        {
        if (approvaltypechkbox.checked)
            {
            aprvltypehiddeninputs[i].style.display = "table"
            }
        else
            {
            aprvltypehiddeninputs[i].style.display = "none"
            }
        }
}

//hide and display light sanity table
function showhidelstable()
{
    var lscheckbox = document.getElementById("lightsanity");
    var lshiddeninputs = document.getElementsByClassName("hideLightSanityTable");
    for (var i=0;i!=lshiddeninputs.length;i++)
        {
        if (lscheckbox.checked)
            {
            lshiddeninputs[i].style.display = "table"
            }
        else
            {
            lshiddeninputs[i].style.display = "none"
            }
        }
}


//to add row in the Approval type sanity table
function addrowforApvlType()
{
	i = 1;
	ApvlTypeincrementor = ApvlTypeincrementor+i;
    var table = document.getElementById("ApprovalTypeTable");
	var rowcount = table.rows.length;

	var row = table.insertRow(rowcount); //inserts at the last
	//var colcount = table.rows[0].cells.length

	//ApvlTypeEnvironment
	var td = document.createElement("td");
	td.setAttribute("align","center");
	var tdselect = document.createElement("select");
	tdselect.name = "ApvlTypeEnvironment"+ApvlTypeincrementor;
	tdselect.id = "ApvlTypeEnvironment"+ApvlTypeincrementor;
	table.appendChild(td);
	td.appendChild(tdselect);
	var ApvlTypeEnvselect = document.getElementById("ApvlTypeEnvironment");
    ApvlTypeEnvselectval = ApvlTypeEnvselect.value;
    tdselect.innerHTML = "<option value="+ApvlTypeEnvselectval+">"+ApvlTypeEnvselectval+"</option>"+
                         '<option value="UAT4">UAT4</option>'+'<option value="UAT1">UAT1</option>'+
                         '<option value="SIT4">SIT4</option>'+'<option value="SIT1">SIT1</option>';
	//tdselect.innerHTML ='<option value="UAT4">UAT4</option>'+

    //
    //tdselect.className = "clsApvlTypeEnvironment";
    row.appendChild(td);

	//ApvlTypescript
	var td = document.createElement("td");
	td.setAttribute("align","center");
	var tdselect = document.createElement("select");
	tdselect.name = "ApvlTypescript"+ApvlTypeincrementor;
	table.appendChild(td);
	td.appendChild(tdselect);
    tdselect.innerHTML = ApvlTypeScriptListLoad();
	row.appendChild(td);

	//lsmachine
	var td = document.createElement("td");
	td.setAttribute("align","center");
	var tdselect = document.createElement("select");
	tdselect.name = "ApvlTypemachine"+ApvlTypeincrementor;
	table.appendChild(td);
	td.appendChild(tdselect);
	tdselect.innerHTML ='<option value="ny2wnlrlg01">ny2wnlrlg01</option>'+
            '<option value="dc1wnlrlg01">dc1wnlrlg01</option>'+
            '<option value="pa2wnlrlg01">pa2wnlrlg01</option>'+
            '<option value="fr2wnlrlg01">fr2wnlrlg01</option>';
	row.appendChild(td);

    //For Execute checkbox field
    //var td = document.createElement("td");
	//td.setAttribute("align","center");
	//var chkbox = document.createElement("input");
	//chkbox.type = "checkbox";
	//chkbox.checked = "checked";
	//chkbox.name = "ApvlTypeexecute"+ApvlTypeincrementor;
	//table.appendChild(td);
    //td.appendChild(chkbox);
	//row.appendChild(td);

	var td = document.createElement("td");
	td.setAttribute("align","center");
	//var status = document.createElement("p");
	table.appendChild(td);
	//td.appendChild(status);
	td.innerHTML = "No Run";
	row.appendChild(td);

	var td = document.createElement("td");
	td.setAttribute("align","center");
	var btn = document.createElement("button");
	btn.onclick = function() {addrowforApvlType();};
	btn.title = "Add Row";
	btn.name = "AddRowApprovalType"
	btn.innerHTML = "+";
	table.appendChild(td);
	td.appendChild(btn);
	row.appendChild(td);

	var td = document.createElement("td");
	td.setAttribute("align","center");
	var btn = document.createElement("button");
	btn.title = "Delete Row";
	btn.name = "Delete Row"
	btn.onclick = function() {deleterowintable(this);};
	btn.innerHTML = "-";
	table.appendChild(td);
	td.appendChild(btn);
	row.appendChild(td);

	document.getElementById("ApprovalTypeScriptCounter").value = ApvlTypeincrementor;

	//var td = document.createElement("td");
	//td.setAttribute("align","center");
	//var btn = document.createElement("button");
	//btn.title = "Stop Script";
	//btn.innerHTML = "X";
	//table.appendChild(td);
	//td.appendChild(btn);
	//row.appendChild(td);
	return false
}



//to add row in the light sanity table
function addrowforlightsanity()
{
	i = 1;
	lsincrementor = lsincrementor+i;
    var table = document.getElementById("LightSanityTable");
	var rowcount = table.rows.length;

	var row = table.insertRow(rowcount); //inserts at the last
	//var colcount = table.rows[0].cells.length

	//lsenvironment
	var td = document.createElement("td");
	td.setAttribute("align","center");
	var tdselect = document.createElement("select");
	tdselect.name = "lsenvironment"+lsincrementor;
	table.appendChild(td);
	td.appendChild(tdselect);
	tdselect.innerHTML ='<option value="UAT4">UAT4</option>'+
                '<option value="SIT4">SIT4</option>'+
                '<option value="SIT1">SIT1</option>';
    row.appendChild(td);

	//lsscript
	var td = document.createElement("td");
	td.setAttribute("align","center");
	var tdselect = document.createElement("select");
	tdselect.name = "lsscript"+lsincrementor;
	table.appendChild(td);
	td.appendChild(tdselect);
	tdselect.innerHTML ='<option value="AP_PrintConfiguredQuote_PCKBP">Print Configured Quote PCKBP</option>'+
	'<option value="Product_Private Cage with kVA Based Power">Private Cage with KVA Based Power</option>'+
	'<option value="Product_ConfigurableAccessories">Configurable Accessories</option>'+
    '<option value="Product_EQX_IEPP_Equinix Internet Exchange Port">Internet Exchange Port</option>';
	row.appendChild(td);

	//lsmachine
	var td = document.createElement("td");
	td.setAttribute("align","center");
	var tdselect = document.createElement("select");
	tdselect.name = "lsmachine"+lsincrementor;
	table.appendChild(td);
	td.appendChild(tdselect);
	tdselect.innerHTML ='<option value="ny2wnlrlg01">ny2wnlrlg01</option>'+
            '<option value="dc1wnlrlg01">dc1wnlrlg01</option>'+
            '<option value="PA2wnlrlg01">PA2wnlrlg01</option>'+
            '<option value="FR2WNLRLG01">FR2WNLRLG01</option>';
	row.appendChild(td);

    var td = document.createElement("td");
	td.setAttribute("align","center");
	var chkbox = document.createElement("input");
	chkbox.type = "checkbox";
	chkbox.checked = "checked";
	chkbox.name = "lsexecute"+lsincrementor;
	table.appendChild(td);
    td.appendChild(chkbox);
	row.appendChild(td);

	var td = document.createElement("td");
	td.setAttribute("align","center");
	var status = document.createElement("p");
	table.appendChild(td);
	td.appendChild(status);
	status.innerHTML = "";
	row.appendChild(td);

	var td = document.createElement("td");
	td.setAttribute("align","center");
	var btn = document.createElement("button");
	btn.onclick = addrowforlightsanity;
	btn.title = "Add Row";
	btn.name = "Add Row"
	btn.innerHTML = "+";
	table.appendChild(td);
	td.appendChild(btn);
	row.appendChild(td);

	var td = document.createElement("td");
	td.setAttribute("align","center");
	var btn = document.createElement("button");
	btn.title = "Delete Row";
	btn.name = "Delete Row"
	btn.onclick = function() {deleterowforlightsanity(this);};
	btn.innerHTML = "-";
	table.appendChild(td);
	td.appendChild(btn);
	row.appendChild(td);

	//var td = document.createElement("td");
	//td.setAttribute("align","center");
	//var btn = document.createElement("button");
	//btn.title = "Stop Script";
	//btn.innerHTML = "X";
	//table.appendChild(td);
	//td.appendChild(btn);
	//row.appendChild(td);
	return false
}

//function to delete the row as needed by the user
function deleterowintable(deletebtn)
{
    var row=deletebtn.parentNode.parentNode;
    row.parentNode.removeChild(row);
}


//get the lsincrementor value + 1
//pass it as value to an Hidden Input box
//from the Input box pass it to the function
//Run through for loop with the range lsincrementor
//read through the field value
//if found then update in the respective xl workbook

//this function would help to set the Environment Row with the value selected at the Main Environment displayed at center of grid
function setRMEnvironmentatRowLevel()
{
    var RMEnvselect = document.getElementById("RMenvironment");
    RMEnvselectval = RMEnvselect.value;
    var RMEnvselect1 = document.getElementById("RMenvironment1");
    RMEnvselect1.value = RMEnvselectval;
    var RMEnvselect2 = document.getElementById("RMenvironment2");
    RMEnvselect2.value = RMEnvselectval;
    var RMEnvselect3 = document.getElementById("RMenvironment3");
    RMEnvselect3.value = RMEnvselectval;
    var RMEnvselect4 = document.getElementById("RMenvironment4");
    RMEnvselect4.value = RMEnvselectval;
}

//Set Environment type for Approval type
function setApvlTypeEnvironmentatRowLevel()
{
    var ApvlTypeEnvselect = document.getElementById("ApvlTypeEnvironment");
    ApvlTypeEnvselectval = ApvlTypeEnvselect.value;
    for (var i = 1;i<=ApvlTypeincrementor;i++)
    {
     var ApvlTypeEnvselectLineItem = document.getElementById("ApvlTypeEnvironment"+i.toString());
     ApvlTypeEnvselectLineItem.value = ApvlTypeEnvselectval;
    }

}

//
function EmailCheck()
{
    emailval = document.getElementById("email");
    Email = emailval.value;
    if (Email=="")
    {
		if (confirm("Press OK, if you would like to continue without Email Id") == true)
        	{
				loadDisplay();	 //css loader
				return true;
            }
           else
                {
                    document.getElementById("email").focus();
                    document.getElementById("email").select();
              		return false;
                }
    }
	loadDisplay();	 //css loader
}

//loading the css loader to display the circle
function loadDisplay() {
  document.getElementById("loader").style.display = "block";
  document.getElementById("loadingText").style.display = "block";
  setTimeout(loadHide,200)
  return true;
}

//calling the loadhide function to trigger after the timeout
function loadHide()
{
  document.getElementById("loader").style.display = "none";
  document.getElementById("loadingText").style.display = "none";
}

//validation to confirm if stop button to be clicked or not
function confirmforStop()
{
        var stopbtn = document.getElementById("Stop");

        if (stopbtn.click)
            {
             	if (confirm("Press OK, if you would like to continue stopping the script execution")== true)
                {
                    disablesendreportandstatuscheck()
                    return true;
                }
                else
                {
                    return false;
                }
            }
}

function disablesendreportandstatuscheck()
{
    document.getElementById("SendReport").disabled = true;
    document.getElementById("Refresh").disabled = true;

}