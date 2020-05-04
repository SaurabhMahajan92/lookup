import * as React from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import { Input, Button, Flex, Grid, Segment, FlexItem, Text, Checkbox, Loader, ButtonProps, CheckboxProps } from '@stardust-ui/react';
import { Anchor } from 'msteams-ui-components-react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { library, IconDefinition } from '@fortawesome/fontawesome-svg-core';
import { faSortAmountUp, faSortAmountDown, faTimesCircle } from '@fortawesome/free-solid-svg-icons';
import "./AddDistributionList.scss";
import { AxiosResponse } from "axios";
import { orderBy } from 'lodash';

export interface IADDistributionList {
    id: string;
    displayName: string
    mail: string
    isSelected: boolean
}

export interface IUserFavoriteDistributionList {
    id: string;
    isPinned: boolean
}

export interface IADDistributionListsProps {
    getADDistributionLists: (query: string) => Promise<AxiosResponse<IADDistributionList[]>>;
    createFavoriteDistributionList: (payload: {}) => Promise<AxiosResponse<void>>;
}

export interface IADDistributionListsState {
    searchResultDistributionLists: IADDistributionList[];
    loader: boolean;
    nameSortIcon: IconDefinition,
    aliasSortIcon: IconDefinition,
    searchQuery: string,
    isHeaderSelected: boolean,
}

//exporting AddDistributionList Component;
export default class AddDistributionList extends React.Component<IADDistributionListsProps, IADDistributionListsState> {

    private skypeMessage: {} = {};

    constructor(props: IADDistributionListsProps) {
        super(props);
        this.state = {
            searchResultDistributionLists: [],
            loader: false,
            nameSortIcon: faSortAmountDown,
            aliasSortIcon: faSortAmountDown,
            searchQuery: "",
            isHeaderSelected: false,
        };

        library.add(faSortAmountUp, faSortAmountDown, faTimesCircle);
    };

    public componentDidMount = () => {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    public componentWillUnmount = () => {
        document.removeEventListener("keydown", this.escFunction, false);
    }

    //Load distribution lists from skype contacts or based on search.
    private dataLoad = () => {
        this.setState({
            loader: true
        });

        //If it is to import distribution lists from skype contacts
        if (this.state.searchQuery) { // If it is based on search
            this.props.getADDistributionLists(this.state.searchQuery).then((response: AxiosResponse<IADDistributionList[]>) => {
                let distributionLists: IADDistributionList[] = [];

                response.data.forEach((currentItem: IADDistributionList) => {
                    distributionLists.push({
                        id: currentItem.id,
                        displayName: currentItem.displayName,
                        mail: currentItem.mail,
                        isSelected: false,
                    });
                });

                this.skypeMessage = "";
                this.setState({
                    searchResultDistributionLists: distributionLists,
                    loader: false
                });

            });
        }
    }

    private onSearchKeyUp = (e: React.KeyboardEvent<HTMLInputElement>) => {
        let searchQuery = (e.target as HTMLInputElement).value;
        this.setState({
            searchQuery: searchQuery
        });
        if (e.keyCode === 13 || (e.key === "Enter")) {
            if (searchQuery) {
                this.dataLoad();
            }
        }
    }

    //To Search data
    private onSearchButtonClick = (e: React.SyntheticEvent<HTMLElement, Event>, v?: ButtonProps) => {
        if (this.state.searchQuery) {
            this.dataLoad();
        }
    }

    //#region Sorting functions

    //Calling appropriate function based on column selected for sorting
    private sortDataByColumn = (column: string, currentIcon: IconDefinition) => {
        if (currentIcon === faSortAmountDown) {
            switch (column) {
                case "displayName":
                    this.sortColumnItems(faSortAmountUp, faSortAmountDown, "displayName", false);
                    break;
                case "mail":
                    this.sortColumnItems(faSortAmountDown, faSortAmountUp, "mail", false);
                    break;
            }
        }
        else {
            this.sortColumnItems(faSortAmountDown, faSortAmountDown, column, true);
        }
    }

    //Setting the sort icons and sorting records
    private sortColumnItems = (faName: IconDefinition, faAlias: IconDefinition, sortColumn: string, sortOrder: boolean) => {

        let distributionLists = this.state.searchResultDistributionLists;
        switch (sortColumn) {
            case "displayName":
                distributionLists = orderBy(distributionLists, [distributionList => distributionList.displayName ? "" : distributionList.displayName.toLowerCase()], sortOrder === true ? ["asc"] : ["desc"]);
                break;
            case "mail":
                distributionLists = orderBy(distributionLists, [distributionList => distributionList.mail ? "" : distributionList.mail.toLowerCase()], sortOrder === true ? ["asc"] : ["desc"]);
                break;
            default:
                break;
        }

        this.setState({
            nameSortIcon: faName,
            aliasSortIcon: faAlias,
            searchResultDistributionLists: distributionLists
        })
    }
    //#endregion

    //When user selected check box, call this function to track checked records
    private onCheckBoxSelect = (e: React.SyntheticEvent<HTMLElement, Event>, checkBoxProps?: CheckboxProps) => {
        let distributionLists: IADDistributionList[] = [];
        let headerCheckBoxSelection = true;
        const selectedChkId = (e.currentTarget as Element).id;

        this.state.searchResultDistributionLists.forEach((currentItem) => {

            if (currentItem.id === selectedChkId) {
                currentItem.isSelected = checkBoxProps!.checked ? checkBoxProps!.checked : false;
            }
            distributionLists.push(currentItem);

            if (!currentItem.isSelected) {
                headerCheckBoxSelection = false;
            }
        });

        this.setState({
            searchResultDistributionLists: distributionLists,
            isHeaderSelected: headerCheckBoxSelection
        });
    }

    //When Select All check box selected
    private onAllCheckBoxSelect = (e: React.SyntheticEvent<HTMLElement, Event>, checkBoxProps?: CheckboxProps) => {
        const headerCheckBoxSelection = checkBoxProps!.checked ? checkBoxProps!.checked : false;
        let distributionLists: IADDistributionList[] = [];
        this.state.searchResultDistributionLists.forEach((currentItem) => {
            currentItem.isSelected = headerCheckBoxSelection
            distributionLists.push(currentItem);
        });

        this.setState({
            searchResultDistributionLists: distributionLists,
            isHeaderSelected: headerCheckBoxSelection
        });
    }

    private escFunction = (e: KeyboardEvent) => {
        if (e.keyCode === 27 || (e.key === "Escape")) {
            microsoftTeams.tasks.submitTask({ "output": "failure" });
        }
    }

    //To add selected distribution lists to favorites.
    private onAddButtonClick = () => {
        let userFavoriteDistributionLists: IUserFavoriteDistributionList[] = [];

        this.state.searchResultDistributionLists.forEach((currentItem) => {
            if (currentItem.isSelected) {
                const userFavoriteDistributionList: IUserFavoriteDistributionList = {
                    id: currentItem.id,
                    isPinned: false,
                };
                userFavoriteDistributionLists.push(userFavoriteDistributionList);
            }
        });

        //Call API to save selected distribution lists to database
        this.postUserFavoriteDistributionLists(userFavoriteDistributionLists).then(() => {
            microsoftTeams.tasks.submitTask({ "output": "success" }); //Close task module on saving
        });
    }

    //Call API to save selected distribution lists to database
    private postUserFavoriteDistributionLists = async (userFavoriteDistributionLists: IUserFavoriteDistributionList[]) => {
        try {
            await this.props.createFavoriteDistributionList(userFavoriteDistributionLists);
        } catch (error) {
            return error;
        }
    }

    public render(): JSX.Element {
        const gridStyle = { width: '100%' };
        if (this.state.loader) {
            return (
                <Loader />
            );
        }
        else {
            const searchResultDistributionLists = this.state.searchResultDistributionLists;
            let segmentRows: {}[] = [];
            searchResultDistributionLists.forEach((currentDL) => {
                segmentRows.push(<Segment className="textAlignCenter">
                    <Checkbox id={currentDL.id} onChange={this.onCheckBoxSelect} checked={currentDL.isSelected} />
                </Segment>);
                segmentRows.push(<Segment content={currentDL.displayName}></Segment>);
                segmentRows.push(<Segment content={currentDL.mail}></Segment>);
            });

            if (this.state.searchResultDistributionLists.length <= 0) {

                return (<div className="taskModule">
                    <div className="formContainer">
                        <Flex gap="gap.small">
                            <FlexItem grow>
                                <Input className="inputField" icon="search" fluid placeholder="Search by Distribution List name" onKeyUp={this.onSearchKeyUp} name="txtSearch" clearable />
                            </FlexItem>
                            <FlexItem push>
                                <Button content="Search" onClick={this.onSearchButtonClick} primary />
                            </FlexItem>
                        </Flex>
                        <div className="formContentContainer" >
                        </div>
                    </div>
                </div>);
            }
            else {

                return (

                    <div className="taskModule">
                        <div className="formContainer">

                            {this.skypeMessage}

                            <Flex gap="gap.small">
                                <FlexItem grow>
                                    <Input icon="search" className="inputField" fluid placeholder="Search by Distribution List name" onKeyUp={this.onSearchKeyUp} name="txtSearch" clearable />
                                </FlexItem>
                                <FlexItem push>
                                    <Button content="Search" onClick={this.onSearchButtonClick} primary />
                                </FlexItem>
                            </Flex>
                            <div className="formContentContainer" >
                                <Grid columns=".5fr 2.5fr 3fr " styles={gridStyle} >
                                    <Segment color="brand" className="textAlignCenter">

                                        <Checkbox onChange={this.onAllCheckBoxSelect} id="chkAll" checked={this.state.isHeaderSelected} />

                                    </Segment>
                                    <Segment color="brand"  >
                                        <Flex gap="gap.small">
                                            <FlexItem grow>
                                                <Text content="Name" />
                                            </FlexItem>
                                            <FlexItem push>
                                                <Anchor href="#" onClick={() => this.sortDataByColumn("displayName", this.state.nameSortIcon)}>
                                                    <FontAwesomeIcon icon={this.state.nameSortIcon} />
                                                </Anchor>
                                            </FlexItem>
                                        </Flex>
                                    </Segment>
                                    <Segment color="brand">
                                        <Flex gap="gap.small">
                                            <FlexItem grow>
                                                <Text content="Alias" />
                                            </FlexItem>
                                            <Anchor href="#" onClick={() => this.sortDataByColumn("mail", this.state.aliasSortIcon)}>
                                                <FlexItem push >
                                                    <FontAwesomeIcon icon={this.state.aliasSortIcon} />
                                                </FlexItem >
                                            </Anchor>
                                        </Flex>
                                    </Segment>

                                    {segmentRows}

                                </Grid>
                            </div>
                            <div className="footerContainer">
                                <div className="buttonContainer">
                                    <Button content="Add" onClick={this.onAddButtonClick} primary className="bottomButton" />
                                </div>
                            </div>

                        </div>
                    </div>
                );
            }

        }
    }
}