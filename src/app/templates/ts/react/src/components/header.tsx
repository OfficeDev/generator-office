import * as React from 'react';

export interface HeaderProps {
    title: string;
    logo: string;
    message: string;
}

export class Header extends React.Component<HeaderProps, any> {
    constructor(props, context) {
        super(props, context);
    }

    render() {
        return (
            <section className='ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500'>
                <img width='90' height='90' src={this.props.logo} alt={this.props.title} title={this.props.title} />
                <h1 className='ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary'>{this.props.message}</h1>
            </section>
        );
    };
};
