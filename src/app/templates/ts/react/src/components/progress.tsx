import * as React from 'react';
import { Spinner, SpinnerType } from 'office-ui-fabric-react';

export interface ProgressProps {
    title: string;
    logo: string;
    message: string;
}

export class Progress extends React.Component<ProgressProps, any> {
    constructor(props, context) {
        super(props, context);
    }

    render() {
        return (
            <section className='ms-welcome__progress ms-u-fadeIn500'>
                <img width='90' height='90' src={this.props.logo} alt={this.props.title} title={this.props.title} />
                <h1 className='ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary'>{this.props.title}</h1>
                <Spinner type={SpinnerType.large} label={this.props.message} />
            </section>
        );
    };
};
