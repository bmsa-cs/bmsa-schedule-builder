/*
    Name:       BMSA Schedule Generator
    Author:     Alex Wolfe
                "What Color Is It?" by Lisa Berry
    Date:       11/26/2017
    Uses scripts located at: https://script.google.com/a/biomedscienceacademy.org/d/1rD9Yo_17dKJIgicvjlWwmc-iTEqfXLWJrvyzkEzSW6pOPJwhCskhVMe2/edit?usp=sharing
*/

function onOpen() {
    //Add a menu to the spreadsheet.
    var ui = SpreadsheetApp.getUi();
    var VERSION = getVersion();
    var VERSION_STRING = "Version: " + VERSION;
    ui.createMenu('BMSA Schedule Generator')
        .addItem('Generate/Update Calendar', 'generateCalendar')
        .addItem('Generate WITHOUT Reminders', 'generateWithoutReminders')
        .addItem('Delete Schedule Calendar', 'deleteCalendar')
        //.addSeparator()
        //.addItem('Check saved personalizations', 'checkPersonalizations')
        //.addItem('Update saved personalizations', 'updatePersonalizations')
        .addSeparator()
        .addItem("What Color Is It?", 'whatColor')
        .addSeparator()
        .addItem("Help", 'showHelp')
        .addItem(VERSION_STRING, 'getVersion')
        .addToUi();
    whatColor();
}

function generateWithoutReminders() {
    generateCalendar("BMSA Schedule", true);
}

function showHelp() {
    showURL("Help", "https://docs.google.com/a/biomedscienceacademy.org/document/d/1HJw_341LT6hBMlrp93w8wp0wojXGimvW8jc7Q9F0PXc/edit?usp=sharing");
}

function showURL(title, href) {
    //var app = UiApp.createApplication().setHeight(50).setWidth(200);
    var html = HtmlService.createHtmlOutput('<a target="_blank" href="' + href + '">' + title + '</a>');
    SpreadsheetApp.getUi()
        .showModalDialog(html, title);
    /*
    app.setTitle(title);
    var link = app.createAnchor('Open BMSA Schedule Generator Help Document', href).setId("link");
    app.add(link);  
    var doc = SpreadsheetApp.getActive();
    doc.show(app);
    */
}