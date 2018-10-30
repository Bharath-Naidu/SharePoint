$(document).ready(function ()
{
    jQuery("#CreateList").click(NewList);
    jQuery("#CreateItem").click(CreateItem);
    jQuery("#CreateField").click(CreateField);
    jQuery("#LINK").click(LINK);
});
function LINK()
{
    //alert("Warning.....");
    var value = $("#textbox2").val();
    var href = "../Lists/" + value;
    $("#LINK").attr('href', href);
}
//function LookUp()
//{
//    var Name = $("#textbox").val();
//    var call = jQuery.ajax({
//        url:_spPageContextInfo.webAbsoluteUrl+"/_api"
//    });
//}


function NewList()
{
    //var Name = $("#textbox").val();
    //var call = jQuery.ajax({
    //    url: _spPageContextInfo.webAbsoluteUrl + "/_api/Web/Lists",
    //    type: "POST",
    //    data: JSON.stringify(
    //        {
    //            "__metadata": { type: "SP.List" },
    //            BaseTemplate: SP.ListTemplateType.tasks,
    //            Title: Name
    //        }),
    //    headers:
    //    {
    //        Accept: "application/json;odata=verbose",
    //        "Content-Type": "application/json;odata=verbose",
    //        "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
    //    }
    //});
    //call.done(function (data, textStatus, jqXHR) {
    //    var message = jQuery("#message");
    //    message.text("List added");
    //});
    //call.fail(function (jqXHR, textStatus, errorThrown) {
    //    var response = JSON.parse(jqXHR.responseText);
    //    var message = response ? response.error.message.value : textStatus;
    //    alert("Call failed. Error: " + message);
    //});
    
}
function CreateItem()
{
    //Retrivedata()
    var call =jQuery.ajax({
        url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/?$select=Title,CurrentUser/Id&$expand=CurrentUser/Id",
        type:"GET",
        dataType: "json",
        headers: {
            Accept:"application/json;odata=verbose"
        }
    });
    call.done(function (data, textStatus, jqXHR) {
        var Userid = data.d.CurrentUser.Id;
        addItem(Userid);
    });
    call.fail(function (data, textStatus, errorThrown) {
        failhandle(data, textStatus, errorThrown);
    });
    function addItem(Userid)
    {
        var Name = $("#textbox1").val();
        var due = new Date();
        due.setDate(due.getDate()+7);
        var call = jQuery.ajax({
            url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/Lists/getByTitle('productsTypes')/Items",
            type: "POST",
            data: JSON.stringify({
                "__metadata": { type: "SP.Data.TasksListItem" },
                Title: Name,
                AssignedToId: Userid,
                DueDate: due
            }),
            headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest":jQuery("#__REQUESTDIGEST").val()
            }
        });
        call.done(function(data, textStatus, jqXHR)  {
            var div = jQuery("#message");
            div.text("Item added");
        });
        call.fail(function (data, textStatus, errorMessgae) {
            failhandle(data,textStatus,errorMessgae);
        });
    }
    function failhandle(data,textStatus,error) {
        var response = JSON.parse(data.responseText);
        var message = response ? response.error.message.value : textStatus;
        alert("Call failed. Error: " + message);
    }
}
//function Retrivedata()
//{
//    $("#textbox").attr("visibility", "visible");
//}

function CreateField() {
    var UserValue = $("#textbox2").val();
    var call = jQuery.ajax({
        url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/Lists/getByTitle('productsTypes')/Fields",
        type:"POST",
        data: JSON.stringify(
            {
                "__metadata": { type: "SP.Field" },
                'FieldTypeKind': 3,
                Title: UserValue
            }),
        headers:
        {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": jQuery("#__REQUESTDIGEST").val()
        }
    });
    call.done(function (data, textStatus, jqXHR) {
        var message = jQuery("#message");
        message.text("Field added");
    });
    call.fail(function (jqXHR, textStatus, errorThrown) {
        var response = JSON.parse(jqXHR.responseText);
        var message = response ? response.error.message.value : textStatus;
        alert("Call failed. Error: " + message);
    });
}









'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage()
{
    var context = SP.ClientContext.get_current();
    var user = context.get_web().get_currentUser();

    // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
    $(document).ready(function () {
        getUserName();
    });

    // This function prepares, loads, and then executes a SharePoint query to get the current users information
    function getUserName() {
        context.load(user);
        context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
    }

    // This function is executed if the above call is successful
    // It replaces the contents of the 'message' element with the user name
    function onGetUserNameSuccess() {
        $('#message').text('Hello ' + user.get_title());
    }

    // This function is executed if the above call fails
    function onGetUserNameFail(sender, args) {
        alert('Failed to get user name. Error:' + args.get_message());
    }
}
