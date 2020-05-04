import * as React from 'react';
import { Header, Menu } from '@stardust-ui/react';
import './HelpPage.scss';
import DLHelp from './DLHelp';
import DLMHelp from './DLMHelp';
import { RouteComponentProps } from 'react-router-dom'

export interface IHelpPageState {
    tabIndex: number;
}

export default class HelpPage extends React.Component<RouteComponentProps, IHelpPageState> {
    constructor(props: RouteComponentProps) {
        super(props);
        this.state = {
            tabIndex: 0
        };
        this.tabClicked = this.tabClicked.bind(this);
    }

    private items: { key: string, content: string }[] = [
        {
            key: "dlheader",
            content: "Distribution Lists Page"
        },
        {
            key: "dlmemberheader",
            content: "Distribution List Members Page"
        }
    ];

    private tabClicked = (item: any, value: any) => {
        if (value.content === "Distribution List Members Page") {
            this.setState({
                tabIndex: 1
            });
        }
        else {
            this.setState({
                tabIndex: 0
            });
        }
    }

    public render(): JSX.Element {
        return (
            <div className="mainComponentHelpPage">
                <div className="header">
                    <Header key="dl"
                        as="h1"
                        color="brand"
                        content="Get Started with DLLookup App"
                    />
                    <Menu items={this.items} defaultActiveIndex={0} underlined primary onItemClick={this.tabClicked} />
                </div>
                {this.state.tabIndex === 0 ? <DLHelp /> : <DLMHelp />}
            </div>
        );
    }
}