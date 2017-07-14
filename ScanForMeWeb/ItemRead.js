/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
https://github.com/OfficeDev/outlook-add-in-command-demo
*/

/// <reference path="/Scripts/jquery.fabric.js" />

(function () {
    'use strict';

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {

            DetermineItemType();
        });
    };

    function DetermineItemType() {
       // var item = Office.cast.item.toItemRead(Office.context.mailbox.item);
        var item1 = Office.context.mailbox.item;
        if (item1.itemType == Office.MailboxEnums.ItemType.Message) {
           var displayText = item1.subject;
            document.getElementById("demo").innerHTML = "this is a message!!!!!!!!";
        }

        if (item1.itemType == Office.MailboxEnums.ItemType.Appointment) {
            var displayText = item1.subject;
            document.getElementById("demo").innerHTML = "this is an appointment"; 
        }
        }
    
})();

// *********************************************************
//
// Outlook-Add-in-ScanForMe, https://github.com/OfficeDev/Outlook-Add-in-ScanForMe
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************