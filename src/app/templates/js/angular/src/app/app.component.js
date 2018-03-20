import { Component } from '@angular/core';
<%- imports %>

import template from './app.component.html';

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';

    run() {
        <%- snippet %>
    }
}