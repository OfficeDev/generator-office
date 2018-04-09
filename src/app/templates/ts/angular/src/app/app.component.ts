import { Component } from '@angular/core';
<%- imports %>

const template = require('./app.component.html');

@Component({
    selector: 'app-home',
    template
})
export default class AppComponent {
    welcomeMessage = 'Welcome';

    async run() {
        <%- snippet %>
    }
}