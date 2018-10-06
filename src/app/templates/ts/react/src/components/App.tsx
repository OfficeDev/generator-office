import * as React from 'react';
import { ButtonType, DefaultButton, Fabric } from 'office-ui-fabric-react';
import Header from './Header';
import HeroList, { HeroListItem } from './HeroList';
import Progress from './Progress';

<%- imports %>

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
    listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
        this.state = {
            listItems: []
        };
    }

    componentDidMount() {
        this.setState({
            listItems: [
                {
                    icon: 'Ribbon',
                    primaryText: 'Achieve more with Office integration'
                },
                {
                    icon: 'Unlock',
                    primaryText: 'Unlock features and functionality'
                },
                {
                    icon: 'Design',
                    primaryText: 'Create and visualize like a pro'
                }
            ]
        });
    }

    click = async () => {
        <%- snippet %>
    }

    render() {
        const {
            title,
            isOfficeInitialized,
        } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Fabric>
                    <Progress
                        title={title}
                        logo='assets/logo-filled.png'
                        message='Please sideload your addin to see app body.'
                    />
                </Fabric>
            );
        }

        return (
            <Fabric>
                <Header logo='assets/logo-filled.png' title={this.props.title} message='Welcome' />
                <HeroList message='Discover what <%= projectDisplayName %> can do for you today!' items={this.state.listItems}>
                    <p className='ms-font-l'>Modify the source files, then click <b>Run</b>.</p>
                    <DefaultButton className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</DefaultButton>
                </HeroList>
            </Fabric>
        );
    }
}
