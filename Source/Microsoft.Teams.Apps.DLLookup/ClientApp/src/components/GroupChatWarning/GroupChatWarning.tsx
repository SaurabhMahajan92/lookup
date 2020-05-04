import * as React from 'react';
import { Button, Flex, ButtonProps } from '@stardust-ui/react';
import * as microsoftTeams from "@microsoft/teams-js";
import './GroupChatWarning.scss';

export interface IGroupChatWarningProps {
    chatListCount: number
}

export default class GroupChatWarning extends React.Component<IGroupChatWarningProps, {}> {
    //#region "Constructor"
    constructor(props: IGroupChatWarningProps) {
        super(props);
        this.onButtonClick = this.onButtonClick.bind(this);
    }

    //#region "React Life Cycle Hooks"
    public componentDidMount = () => {
        microsoftTeams.initialize();
    }
    //#endregion

    //#region "On Button Click"
    private onButtonClick = (e: React.SyntheticEvent<HTMLElement, Event>, v?: ButtonProps) => {
        microsoftTeams.tasks.submitTask({ "response": (e.currentTarget as Element).id });
    }
    //#endregion

    public render(): JSX.Element {
        let styles = { padding: '5%' };
        return (
            <div style={styles}>
                <Flex>
                    <div>
                        <p>We currently support group chat with 8 users in Teams and you have selected to start a chat with {this.props.chatListCount} members in this distribution list.</p>

                        <p>To move forward, would you like to chat with most recent 8 members available online?</p>

                    </div>
                </Flex>
                <div className="footerContainer">
                    <div className="buttonContainer">
                        <Button key="accept" id="YES" value="YES" primary onClick={this.onButtonClick}>Yes</Button> <Button key="decline" id="NO" value="NO" primary onClick={this.onButtonClick}>No</Button>
                    </div>
                </div>
            </div>
        );
    }
}
