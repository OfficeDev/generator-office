/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

<%- imports %>

$(document).ready(() => {
  $('#run').click(run);
});

// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
  $('#sideload-msg').hide();
  $('#app-body').show();
};

async function run() {
  <%- snippet %>
}

