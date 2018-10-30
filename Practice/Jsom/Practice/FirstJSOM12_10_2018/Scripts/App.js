$(document).ready(function () {
    jQuery("#Create").click(createList);
    jQuery("#Display").click(DisplayAllLists);
    jQuery("#DisplayList").click(DisplayList);
    jQuery("#enterdata").click(InsertData);
});

function InsertData()
{
    alert("read to insert");
    var Title = document.getElementById("takevalue").value;
    var context = SP.ClientContext.get_current();
    var web = context.get_web();
    
    try
    {
        var list = web.get_lists().getByTitle("NewList");
        var listItem = new SP.ListCreationInformation();
        var item = list.addItem(listItem);
        item.set_item("Title", Title);
        item.set_item("AssignedTo", web.get_currentUser());
        var date = new Date();
        date.setDate(date.getDate() + 7);
        item.set_item("DueDate", date);
        item.update();
        context.executeQueryAsync(Sucess,fail);
    }
    catch (ex)
    {
        alert(ex.message);
    }
    function Sucess()
    {
        var me = jQuery("#message");
        me.text("Succesfully inserted data");
    }
    function fail()
    {
        var me = jQuery("#message");
        me.text("Something Wrong..........");
    }
    alert("finish");
}




function DisplayList() {
    alert("Display now");
    var context = SP.ClientContext.get_current();
    var web = context.get_web();
    var lists = web.get_lists();
    context.load(web,"Title","Description");
    context.load(lists,"Include(Title)");
    context.executeQueryAsync(success, fail);
    function success()
    {
        var me = jQuery("#message");
        me.text(web.get_title());
        var Lists = lists.getEnumerator();
        while (Lists.moveNext())
        {
            me.append("<br/>")
            me.append(Lists.get_current().get_title());
        }
    }
    function fail()
    {
        me.text("Sorry Something Wrong.............");
    }
    alert("here");
}

function DisplayAllLists()
{
    alert("working");
    var context = SP.ClientContext.get_current();
    var web = context.get_web();
    var lists = web.get_lists();
    context.load(web, "Title", "Description");
    context.load(lists, "Include(Title, Fields.Include(Title))");
    context.executeQueryAsync(success, fail);

    function success()
    {
        var message = jQuery("#message");
        message.text(web.get_title());
        var lenum = lists.getEnumerator();
        while (lenum.moveNext())
        {
            var list = lenum.get_current();
            message.append("<br />");
            message.append(list.get_title());
            var fenum = list.get_fields().getEnumerator();
            var i = 0;
            while (fenum.moveNext())
            {
                var field = fenum.get_current();
                message.append("<br />&nbsp;&nbsp;&nbsp;&nbsp;");
                message.append(field.get_title());
               // if (i++ > 5) break;
            }
        }
    }
        
    function fail(sender, args) {
        alert("Call failed. Error: " +
            args.get_message());
    }
    alert("working Fine");
}

function createList() {
    var context = SP.ClientContext.get_current();
    var web = context.get_web();

    try {
        var lci = new SP.ListCreationInformation();
        lci.set_title("NewList");
        lci.set_templateType(SP.ListTemplateType.tasks);
        lci.set_quickLaunchOption(SP.QuickLaunchOptions.on);
        var list = web.get_lists().add(lci);

        context.executeQueryAsync(success, fail);
    } catch (ex) {
        alert(ex.message);
    }

    function success() {
        var message = jQuery("#message");
        message.text("List added");
    }

    function fail(sender, args) {
        alert("Call failed. Error: " +
            args.get_message());
    }
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
