import * as React from 'react';
import { BrowserRouter, Route, Switch } from 'react-router-dom';
import './App.scss';
import { Provider, themes } from '@stardust-ui/react';
import * as microsoftTeams from "@microsoft/teams-js";
import { TeamsThemeContext, getContext, ThemeStyle } from 'msteams-ui-components-react';
import ErrorPage from "./components/ErrorPage/errorPage";
import SignInPage from "./components/SignInPage/signInPage";
import SignInSimpleStart from "./components/SignInPage/signInSimpleStart";
import SignInSimpleEnd from "./components/SignInPage/signInSimpleEnd";
import DistributionLists from './components/DistributionLists/DistributionLists';
import AddDistributionList from './components/AddDistributionList/AddDistributionList';
import DistributionListMembers from './components/DistributionListMembers/DistributionListMembers';
import GroupChatWarning from './components/GroupChatWarning/GroupChatWarning';
import HelpPage from './components/HelpPage/HelpPage';
import { createFavoriteDistributionList, getADDistributionLists, pinStatusUpdate, getDistributionListsMembers, getFavoriteDistributionLists, getDistributionListMembersOnlineCount, getUserPresence, getUserPageSizeChoice, createUserPageSizeChoice, getClientId } from './apis/apiList';

export interface IAppState {
    theme: string;
    themeStyle: number;
}

class App extends React.Component<{}, IAppState> {

    constructor(props: {}) {
        super(props);
        this.state = {
            theme: "",
            themeStyle: ThemeStyle.Light,
        }
    }

    public componentDidMount = () => {
        microsoftTeams.initialize();
        microsoftTeams.getContext((context) => {
            let theme = context.theme || "";
            this.updateTheme(theme);
            this.setState({
                theme: theme
            });
        });

        microsoftTeams.registerOnThemeChangeHandler((theme) => {
            this.updateTheme(theme);
            this.setState({
                theme: theme,
            }, () => {
                this.forceUpdate();
            });
        });
    }

    public setThemeComponent = () => {
        if (this.state.theme === "dark") {
            return (
                <Provider theme={themes.teamsDark}>
                    <div className="darkContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        }
        else if (this.state.theme === "contrast") {
            return (
                <Provider theme={themes.teamsHighContrast}>
                    <div className="highContrastContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        } else {
            return (
                <Provider theme={themes.teams}>
                    <div className="defaultContainer">
                        {this.getAppDom()}
                    </div>
                </Provider>
            );
        }
    }

    private updateTheme = (theme: string) => {
        if (theme === "dark") {
            this.setState({
                themeStyle: ThemeStyle.Dark
            });
        } else if (theme === "contrast") {
            this.setState({
                themeStyle: ThemeStyle.HighContrast
            });
        } else {
            this.setState({
                themeStyle: ThemeStyle.Light
            });
        }
    }

    public getAppDom = () => {
        const context = getContext({
            baseFontSize: 10,
            style: this.state.themeStyle
        });
        return (
            <TeamsThemeContext.Provider value={context}>
                <div className="appContainer">
                    <BrowserRouter>
                        <Switch>
                            <Route exact path="/dls" render={(props) => <DistributionLists {...props} getFavoriteDistributionLists={getFavoriteDistributionLists} getDistributionListMembersOnlineCount={getDistributionListMembersOnlineCount}  getUserPageSizeChoice={getUserPageSizeChoice} createUserPageSizeChoice={createUserPageSizeChoice} getClientId={getClientId} />} />
                            <Route exact path="/dlmemberlist/:id/:name" render={(props) => <DistributionListMembers {...props} parentDLID={props.match.params.id} parentDLName={props.match.params.name} getDistributionListsMembers={getDistributionListsMembers} pinStatusUpdate={pinStatusUpdate} getUserPresence={getUserPresence} getUserPageSizeChoice={getUserPageSizeChoice} createUserPageSizeChoice={createUserPageSizeChoice} />} />
                            <Route exact path="/adfavorite/:isskypedl?" render={(props) => <AddDistributionList {...props} getADDistributionLists={getADDistributionLists} createFavoriteDistributionList={createFavoriteDistributionList} />} />
                            <Route exact path="/groupchatwarning/:count" render={(props) => <GroupChatWarning {...props} chatListCount={props.match.params.count} />} />
                            <Route exact path="/errorpage" component={ErrorPage} />
                            <Route exact path="/errorpage/:id" component={ErrorPage} />
                            <Route exact path="/signin" component={SignInPage} />
                            <Route exact path="/signin-simple-start" component={SignInSimpleStart} />
                            <Route exact path="/signin-simple-end" component={SignInSimpleEnd} />
                            <Route exact path="/help" component={HelpPage} />
                        </Switch>
                    </BrowserRouter>
                </div>
            </TeamsThemeContext.Provider>
        );
    }

    public render(): JSX.Element {
        return (
            <div>
                {this.setThemeComponent()}
            </div>
        );
    }
}

export default App;