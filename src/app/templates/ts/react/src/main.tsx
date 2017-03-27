import * as React from 'react';
import { render } from 'react-dom';
import { App } from './components/app';
import './assets/styles/global.scss';

function main() {
    render(
        <App title='<%= projectDisplayName %>' />,
        document.querySelector('#container')
    );
}

main();
