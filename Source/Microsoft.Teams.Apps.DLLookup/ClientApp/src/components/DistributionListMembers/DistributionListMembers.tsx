import * as React from 'react';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { getBaseUrl } from '../../configVariables';
import * as microsoftTeams from "@microsoft/teams-js";
import { Loader, Flex, Text, Segment, FlexItem, Checkbox, Button, Grid, Input, Dropdown, DropdownProps, CheckboxProps } from '@stardust-ui/react';
import { Anchor } from 'msteams-ui-components-react';
import { IconDefinition } from '@fortawesome/fontawesome-svg-core';
import { faSortAmountUp, faSortAmountDown, faCheckCircle, faCircle, faMinusCircle, faClock } from '@fortawesome/free-solid-svg-icons';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import Pagination from '../Pagination/Pagination';
import './DistributionListMembers.scss';
import { chunk } from 'lodash';
import { AxiosResponse } from "axios";
import { orderBy } from 'lodash';

export interface ITaskInfo {
    title?: string;
    height: number;
    width: number;
    url: string;
    fallbackUrl: string;
}

export interface IDistributionListMember {
    id: string;
    displayName: string;
    jobTitle: string;
    mail: string;
    userPrincipalName: string;
    isPinned: boolean;
    presence: string;
    isSelected: boolean;
    isGroup: boolean;
    sortOrder: number;
    type: string;
}

export interface IPresenceData {
    userPrincipalName: string;
    availability: string;
    availabilitySortOrder: number;
    id: string;
}

export interface IUserPageSizeChoice {
    distributionListPageSize: number;
    distributionListMemberPageSize: number;
}

export interface IDistributionListMembersProps {
    parentDLID: string;
    parentDLName: string;
    getDistributionListsMembers: (groupID?: string) => Promise<AxiosResponse<IDistributionListMember[]>>;
    pinStatusUpdate: (pinnedUser: string, status: boolean, distributionListID: string) => Promise<AxiosResponse<void>>;
    getUserPresence: (payload: {}) => Promise<AxiosResponse<IPresenceData[]>>;
    createUserPageSizeChoice: (payload: {}) => Promise<AxiosResponse<void>>;
    getUserPageSizeChoice: () => Promise<AxiosResponse<IUserPageSizeChoice>>;
}

export interface IDistributionListMembersState {
    distributionListMembers: IDistributionListMember[];
    loader: boolean;
    activePage: number;
    nameSortIcon: IconDefinition;
    aliasSortIcon: IconDefinition;
    titleSortIcon: IconDefinition;
    presenceSortIcon: IconDefinition;
    masterDistributionListMembers: IDistributionListMember[];      //Copy of DL MemberList
    isAllSelectChecked: boolean;
    pageSize: number;
    isGoBackClicked: boolean;
}

//Exporting DistributionListMembers component
export default class DistributionListMembers extends React.Component<IDistributionListMembersProps, IDistributionListMembersState> {

    private isOpenTaskModuleAllowed: boolean;
    private checkedMembersForChat: IDistributionListMember[];
    private historyArray: string[];
    private batchRequestLimit: number = 40;
    private groupChatMembersLimit: number = 8;
    private defaultPageSize: number = 20;
    private notYetFetchedText: string = "Not yet fetched";
    private readonly taskModulePositiveResponseString: string = "YES";
    private readonly availabilityStatusOnline: string = "Online";
    private readonly pageId: number = 2; //DistributionListMembers.tsx treating as Page id 2

    constructor(props: IDistributionListMembersProps) {
        super(props);
        initializeIcons();
        this.isOpenTaskModuleAllowed = true;
        this.checkedMembersForChat = [];
        this.historyArray = [];
        this.state = {
            distributionListMembers: [],
            loader: true,
            activePage: 0,
            nameSortIcon: faSortAmountDown,
            aliasSortIcon: faSortAmountDown,
            titleSortIcon: faSortAmountDown,
            presenceSortIcon: faSortAmountDown,
            masterDistributionListMembers: [],
            isAllSelectChecked: false,
            pageSize: this.defaultPageSize,
            isGoBackClicked: false
        };
        this.checkboxChanged = this.checkboxChanged.bind(this);
        this.selectAllCheckboxChanged = this.selectAllCheckboxChanged.bind(this);
        this.pinStatusUpdate = this.pinStatusUpdate.bind(this);
        this.groupChatWithMembers = this.groupChatWithMembers.bind(this);
        this.oneOnOneChat = this.oneOnOneChat.bind(this);
    }

    public componentDidMount = () => {
        const historyJson = localStorage.getItem("localStorageHistory");
        if (historyJson != null) {
            this.historyArray = JSON.parse(historyJson);
            this.historyArray.push(window.location.href);
            localStorage.setItem("localStorageHistory", JSON.stringify(this.historyArray));
        }
        else {
            this.historyArray.push(window.location.href);
            localStorage.setItem("localStorageHistory", JSON.stringify(this.historyArray));
        }

        this.getPageSize();

        this.dataLoad();
        this.resetSorting(this.state.distributionListMembers);
    }

    //This function is to load data to state using API.
    private dataLoad = () => {

        //API call to get the members of group
        this.props.getDistributionListsMembers(this.props.parentDLID).then((response: AxiosResponse<IDistributionListMember[]>) => {
            const members = response.data;
            let distributionListMembersTemp: IDistributionListMember[] = [];
            for (let i = 0; i < members.length; i++) {
                distributionListMembersTemp.push(
                    {
                        id: members[i].id,
                        displayName: members[i].displayName,
                        jobTitle: members[i].jobTitle === null ? "" : members[i].jobTitle,
                        userPrincipalName: members[i].userPrincipalName,
                        mail: members[i].mail,
                        presence: (members[i].type === "#microsoft.graph.group") ? "" : this.notYetFetchedText,
                        isPinned: members[i].isPinned,
                        isSelected: false,
                        isGroup: members[i].type === "#microsoft.graph.group",
                        sortOrder: 10,//Any number greater than 5 is fine,
                        type: members[i].type
                    }
                );
            }
            this.resetSorting(distributionListMembersTemp);
            this.getAllUserPresenceAsync();
            this.setState({
                loader: false
            });
        });
    }

    //To get group members presence information
    private getAllUserPresenceAsync = async () => {
        let presenceDataList: IPresenceData[] = [];

        this.state.masterDistributionListMembers.forEach((currentDistributionListMember) => {

            if (currentDistributionListMember.presence === this.notYetFetchedText) {
                presenceDataList.push({
                    userPrincipalName: currentDistributionListMember.userPrincipalName,
                    availability: "",
                    availabilitySortOrder: 0,
                    id: currentDistributionListMember.id
                });
            }
        });

        let batchRequests = chunk(presenceDataList, this.batchRequestLimit);
        for (let i = 0; i < batchRequests.length; i++) {
            this.getUserPresenceAsync(batchRequests[i]);
        }
    }

    //To get user presence
    private getUserPresenceAsync = async (iPresenceDataList: IPresenceData[]) => {
        this.props.getUserPresence(iPresenceDataList).then((response: AxiosResponse<IPresenceData[]>) => {
            const presenceDataList: IPresenceData[] = response.data;

            //Set the state for user presence in master distribution list
            const masterDistributionListMembers = this.state.masterDistributionListMembers.map((currentItem) => {
                if (currentItem.userPrincipalName != null) {
                    let presenceDetailsOfCurrentItem = presenceDataList.find((currentPresenceRecord: IPresenceData) => currentPresenceRecord.userPrincipalName.toLowerCase() === currentItem.userPrincipalName.toLowerCase());
                    if (presenceDetailsOfCurrentItem !== undefined) {
                        currentItem.presence = presenceDetailsOfCurrentItem.availability;
                        currentItem.sortOrder = presenceDetailsOfCurrentItem.availabilitySortOrder;
                        currentItem.id = presenceDetailsOfCurrentItem.id;
                    }
                }
                return currentItem;
            });

            //Set the state for user presence in distribution list
            const distributionListMembers = this.state.distributionListMembers.map((currentItem) => {
                if (currentItem.userPrincipalName != null) {
                    let presenceDetailsOfCurrentItem = presenceDataList.find((currentPresenceRecord: IPresenceData) => currentPresenceRecord.userPrincipalName.toLowerCase() === currentItem.userPrincipalName.toLowerCase());
                    if (presenceDetailsOfCurrentItem !== undefined) {
                        currentItem.presence = presenceDetailsOfCurrentItem.availability;
                        currentItem.sortOrder = presenceDetailsOfCurrentItem.availabilitySortOrder;
                        currentItem.id = presenceDetailsOfCurrentItem.id;
                    }
                }
                return currentItem;
            });

            this.setState({
                masterDistributionListMembers: masterDistributionListMembers,
                distributionListMembers: distributionListMembers,
            })

            this.sortColumnItems(faSortAmountDown, faSortAmountDown, faSortAmountDown, faSortAmountDown, faSortAmountDown, "presence", true, "sortOrder");
        });
    }

    // "Render Corresponding Presence Icon"
    private renderPresenceInfo = (presence: string) => {
        switch (presence) {
            case "None":
                return {
                    "icon": faCircle,
                    "color": "#D3D3D3"
                };
            case "Away":
                return {
                    "icon": faClock,
                    "color": "#FDB913"
                };
            case "Offline":
                return {
                    "icon": faCircle,
                    "color": "#D3D3D3"
                };
            case "DoNotDisturb":
                return {
                    "icon": faMinusCircle,
                    "color": "#C4314B"
                };
            case "BeRightBack":
                return {
                    "icon": faClock,
                    "color": "#FDB913"
                };

            case "Busy":
                return {
                    "icon": faCircle,
                    "color": "#C4314B"
                };
            case "Online":
                return {
                    "icon": faCheckCircle,
                    "color": "#92C353"
                };
            default:
                return {
                    "icon": faCircle,
                    "color": "#D3D3D3"
                };
        }

    }

    //#region "Sorting functions"

    //Calling appropriate function based on column selected for sorting
    private sortDataByColumn = (column: string, currentIcon: IconDefinition) => {
        if (currentIcon === faSortAmountDown) {
            switch (column) {
                case "displayName":
                    this.sortColumnItems(faSortAmountUp, faSortAmountDown, faSortAmountDown, faSortAmountDown, faSortAmountDown, column, false);
                    break;
                case "mail":
                    this.sortColumnItems(faSortAmountDown, faSortAmountUp, faSortAmountDown, faSortAmountDown, faSortAmountDown, column, false);
                    break;
                case "presence":
                    this.sortColumnItems(faSortAmountDown, faSortAmountDown, faSortAmountDown, faSortAmountUp, faSortAmountDown, column, false, "sortOrder");
                    break;
                case "lastSeen":
                    this.sortColumnItems(faSortAmountDown, faSortAmountDown, faSortAmountDown, faSortAmountDown, faSortAmountUp, column, false, "lastSeenTime");
                    break;
            }
        }
        else {
            let sortColumn = column;
            if (column === "presence") {
                sortColumn = "sortOrder";
            }
            this.sortColumnItems(faSortAmountDown, faSortAmountDown, faSortAmountDown, faSortAmountDown, faSortAmountDown, column, true, sortColumn);
        }
    }

    //Setting the sort icons and sorting pinned-unpinned records separately
    private sortColumnItems = (faName: IconDefinition, faAlias: IconDefinition, faTitle: IconDefinition, faPresence: IconDefinition, faLastSeen: IconDefinition, sortColumn: string, sortOrder: boolean, sortDataColumn?: string) => {

        let distributionListMembers = this.state.distributionListMembers;
        let pinnedRecords = distributionListMembers.filter((e: IDistributionListMember) => e.isPinned === true);
        let unpinnedRecords = distributionListMembers.filter((e: IDistributionListMember) => e.isPinned === false);

        switch (sortColumn) {
            case "displayName":
                pinnedRecords = orderBy(pinnedRecords, [pinnedRecord => pinnedRecord.displayName.toLowerCase()], sortOrder === true ? ["asc"] : ["desc"]);
                unpinnedRecords = orderBy(unpinnedRecords, [unpinnedRecord => unpinnedRecord.displayName.toLowerCase()], sortOrder === true ? ["asc"] : ["desc"]);
                break;
            case "mail":
                pinnedRecords = orderBy(pinnedRecords, [pinnedRecord => pinnedRecord.mail ? "" : pinnedRecord.mail.toLowerCase()], sortOrder === true ? ["asc"] : ["desc"]);
                unpinnedRecords = orderBy(unpinnedRecords, [unpinnedRecord => unpinnedRecord.mail ? "" : unpinnedRecord.mail.toLowerCase()], sortOrder === true ? ["asc"] : ["desc"]);
                break;
            case "presence":
                pinnedRecords = orderBy(pinnedRecords, ["sortOrder"], sortOrder === true ? ["asc"] : ["desc"]);
                unpinnedRecords = orderBy(unpinnedRecords, ["sortOrder"], sortOrder === true ? ["asc"] : ["desc"]);
                break;
            default:
                break;
        }

        distributionListMembers = pinnedRecords.concat(unpinnedRecords);
        this.setState({
            nameSortIcon: faName,
            aliasSortIcon: faAlias,
            titleSortIcon: faTitle,
            presenceSortIcon: faPresence,
            distributionListMembers: distributionListMembers
        })
    }

    //Used to reset the sorting on data load
    private resetSorting = (distributionListMembers: IDistributionListMember[]) => {
        let pinnedRecords = distributionListMembers.filter((e: IDistributionListMember) => e.isPinned === true);
        let unpinnedRecords = distributionListMembers.filter((e: IDistributionListMember) => e.isPinned === false);
        let sortColumn = "presence";
        let sortOrder = false;

        if (this.state.nameSortIcon === faSortAmountUp) {
            sortColumn = "displayName";
            sortOrder = false;
        }
        else if (this.state.aliasSortIcon === faSortAmountUp) {
            sortColumn = "mail";
            sortOrder = false;
        }
        else if (this.state.titleSortIcon === faSortAmountUp) {
            sortColumn = "title";
            sortOrder = false;
        }
        else if (this.state.presenceSortIcon === faSortAmountUp) {
            sortColumn = "presence";
            sortOrder = false;
        }

        pinnedRecords = orderBy(pinnedRecords, sortColumn, ["desc"]);
        unpinnedRecords = orderBy(unpinnedRecords, sortColumn, ["desc"]);
        distributionListMembers = pinnedRecords.concat(unpinnedRecords);

        this.setState({
            distributionListMembers: distributionListMembers,
            masterDistributionListMembers: distributionListMembers
        });
    }
    //#endregion "Sorting functions"

    //"Search function"
    private search = (e: React.SyntheticEvent<HTMLElement, Event>) => {
        let searchQuery = (e.target as HTMLInputElement).value;
        if (!searchQuery) {
            this.setState({
                distributionListMembers: this.state.masterDistributionListMembers,
            })
        }
        else {
            this.setState({
                distributionListMembers: this.state.masterDistributionListMembers.filter((x: IDistributionListMember) => x.displayName.toLowerCase().includes(searchQuery.toLowerCase())),
                activePage: 0,
            })
        }
    }

    // "Individual record checkbox selected"
    private checkboxChanged = (e: React.SyntheticEvent<HTMLElement, Event>, v?: CheckboxProps) => {
        let headerCheckValue = true;
        const selectedChkId = (e.currentTarget as Element).id;
        this.state.distributionListMembers.forEach((currentItem) => {
            if (currentItem.id === selectedChkId) {
                currentItem.isSelected = v!.checked ? v!.checked : false;
                if (currentItem.isSelected) {
                    this.checkedMembersForChat.push(currentItem);
                }
                else {
                    this.checkedMembersForChat.splice(this.checkedMembersForChat.findIndex(item => item.userPrincipalName === currentItem.userPrincipalName), 1);
                }
            }

            if (!currentItem.isSelected) {
                headerCheckValue = false;
            }
        });

        this.setState({
            isAllSelectChecked: headerCheckValue
        });
    }

    // "All Select Checkbox selected"
    private selectAllCheckboxChanged = (e: React.SyntheticEvent<HTMLElement, Event>, v?: CheckboxProps) => {
        const headerChkValue = v!.checked ? v!.checked : false;
        if (headerChkValue) {
            this.state.distributionListMembers.forEach((currentItem) => {
                if (!currentItem.isGroup) {
                    currentItem.isSelected = headerChkValue;
                    this.checkedMembersForChat.push(currentItem);
                }
            });
            this.setState({
                isAllSelectChecked: headerChkValue
            });
        }
        else {
            this.state.distributionListMembers.forEach((currentItem) => {
                currentItem.isSelected = headerChkValue;
            });
            this.checkedMembersForChat = [];
            this.setState({
                isAllSelectChecked: headerChkValue
            });
        }
    }

    //To update pin status
    private pinStatusUpdate = (e: React.MouseEvent<HTMLAnchorElement, MouseEvent>) => {
        const pinId = (e.target as Element).id;
        const member = (this.state.distributionListMembers.filter((x: IDistributionListMember) => { return x.id === pinId }));
        const pinStatus = !member[0].isPinned;

        //API call to update the database depending on whether the user pinned or not
        this.props.pinStatusUpdate(pinId, pinStatus, this.props.parentDLID).then((response: AxiosResponse<void>) => {
            this.state.distributionListMembers.forEach((x: IDistributionListMember) => {
                if (pinId === x.id) {
                    x.isPinned = pinStatus;
                }
            });
            this.state.masterDistributionListMembers.forEach((x: IDistributionListMember) => {
                if (pinId === x.id) {
                    x.isPinned = pinStatus;
                }
            });

            this.setState({
                distributionListMembers: this.state.distributionListMembers,
                masterDistributionListMembers: this.state.masterDistributionListMembers
            })
            this.resetSorting(this.state.distributionListMembers);
        });
    }

    //#region "Set Current Page for Pagination"
    private setActivePage = (newPageNumber: number) => {
        this.setState({
            activePage: newPageNumber,
        })
    }

    // "Helper for groupChat"
    private groupChatLink = () => {
        let userList = this.checkedMembersForChat.map(members => members.userPrincipalName).join(',');
        return userList;
    }

    // "groupChat from Chat for Nested DL"
    private groupChatWithMembers = () => {
        if (this.checkedMembersForChat.length > this.groupChatMembersLimit) {
            this.onOpenTaskModule();
        }
        else {
            const url = "https://teams.microsoft.com/l/chat/0/0?users=" + this.groupChatLink();
            microsoftTeams.executeDeepLink(encodeURI(url));
        }
    }

    //"1 on 1 Chat"
    private oneOnOneChat = (e: React.MouseEvent<HTMLAnchorElement, MouseEvent>) => {
        const url = "https://teams.microsoft.com/l/chat/0/0?users=" + encodeURI((e.target as Element).id);
        microsoftTeams.executeDeepLink(encodeURI(url));
    }

    //Action to navigate back to previous page
    private pageGoBack = () => {
        this.setState({
            isGoBackClicked: true
        });
        const historyJson = localStorage.getItem("localStorageHistory");
        if (historyJson != null) {
            this.historyArray = JSON.parse(historyJson);
            const goToPage = this.historyArray[this.historyArray.length - 2];
            this.historyArray.splice(this.historyArray.length - 2, 2);
            localStorage.setItem("localStorageHistory", JSON.stringify(this.historyArray));
            window.location.href = goToPage;
        }
    }

    //Get Page size from database/local storage
    private getPageSize = async () => {
        if (localStorage.getItem('localStorageDLPageSizeValue') === null || localStorage.getItem('localStorageDLPageSizeValue') === undefined) {
            this.props.getUserPageSizeChoice().then((response: AxiosResponse<IUserPageSizeChoice>) => {
                if (response.data) {
                    this.setState({
                        pageSize: response.data.distributionListPageSize === 0 ? this.defaultPageSize : response.data.distributionListPageSize
                    });
                    localStorage.setItem('localStorageDLPageSizeValue', response.data.distributionListPageSize === 0 ? this.defaultPageSize.toString() : response.data.distributionListPageSize.toString());
                    localStorage.setItem('localStorageDLMembersPageSizeValue', response.data.distributionListMemberPageSize === 0 ? this.defaultPageSize.toString() : response.data.distributionListMemberPageSize.toString());
                }
                else {
                    localStorage.setItem('localStorageDLPageSizeValue', this.defaultPageSize.toString());
                    localStorage.setItem('localStorageDLMembersPageSizeValue', this.defaultPageSize.toString());
                }

            });
        }
        else {
            this.setState({
                pageSize: Number(localStorage.getItem('localStorageDLPageSizeValue'))
            });
        }
    }

    //setting page size
    private setPageSize = (e: React.SyntheticEvent<HTMLElement, Event>, v?: DropdownProps) => {
        this.setState({
            pageSize: Number(v!.value),
            activePage: 0,
        });

        //Update database
        this.props.createUserPageSizeChoice({
            "PageId": this.pageId,
            "PageSize": v!.value
        }).then((response: AxiosResponse<void>) => {
            localStorage.setItem('localStorageDLMembersPageSizeValue', (v!.value || this.defaultPageSize).toString());
        })
    }

    //"Group chat task module"
    private onOpenTaskModule = () => {
        if (this.isOpenTaskModuleAllowed) {
            this.isOpenTaskModuleAllowed = false;
            const taskInfo: ITaskInfo = {
                url: getBaseUrl() + "/groupchatwarning/" + this.checkedMembersForChat.length,
                title: "",
                height: 300,
                width: 400,
                fallbackUrl: getBaseUrl() + "/groupchatwarning" + this.checkedMembersForChat.length
            }

            const submitHandler = (err: string, result: any) => {
                this.isOpenTaskModuleAllowed = true;
                if (result.response === this.taskModulePositiveResponseString) {
                    this.checkedMembersForChat = this.checkedMembersForChat.filter(item => item.presence === this.availabilityStatusOnline);
                    if (this.checkedMembersForChat.length > this.groupChatMembersLimit) {
                        this.checkedMembersForChat.splice(this.groupChatMembersLimit, this.checkedMembersForChat.length);
                    }
                    this.groupChatWithMembers();
                }
            };

            microsoftTeams.tasks.startTask(taskInfo, submitHandler);
        }
    }

    //"Render Method"
    public render(): JSX.Element {
        //Page size drop down values.
        let pageSize = [20, 50, 100];
        let pageNumber: number = this.state.activePage;
        let index = pageSize.indexOf(this.state.pageSize);
        let items = []; //Populate grid

        for (let j: number = pageNumber * this.state.pageSize; j < (pageNumber * this.state.pageSize) + this.state.pageSize; j++) {
            //#region Populate Grid
            if (j >= this.state.distributionListMembers.length) {
                break;
            }
            const distributionListMember = this.state.distributionListMembers[j];

            items.push(<Segment >
                <Flex gap="gap.small">
                    <FlexItem grow>
                        <Checkbox key={distributionListMember.userPrincipalName} id={distributionListMember.id} label={distributionListMember.displayName} onClick={this.checkboxChanged} checked={distributionListMember.isSelected} disabled={distributionListMember.isGroup} />
                    </FlexItem>
                    <FlexItem>
                        <Icon iconName="Pinned" hidden={!distributionListMember.isPinned} />
                    </FlexItem>
                </Flex>
            </Segment>);

            items.push(<Segment content={distributionListMember.mail} ></Segment>);

            if (this.state.distributionListMembers[j].presence === this.notYetFetchedText) {
                items.push(<Segment><Loader size="smallest" /></Segment>)
            }
            else if (this.state.distributionListMembers[j].presence === "") {
                items.push(
                    <Segment>
                        <Flex gap="gap.small">
                        </Flex>
                    </Segment>);
            }
            else {
                const userPresence = this.renderPresenceInfo(this.state.distributionListMembers[j].presence);
                items.push(
                    <Segment>
                        <Flex gap="gap.small">
                            <FlexItem><FontAwesomeIcon icon={userPresence.icon} style={{ color: userPresence.color }} /></FlexItem>
                            <FlexItem><Text content={this.state.distributionListMembers[j].presence} /></FlexItem>
                        </Flex>
                    </Segment>);

            }

            if (distributionListMember.isGroup) {
                items.push(<Segment >
                    <Flex gap="gap.small">
                        <Anchor href={"/dlmemberlist/" + distributionListMember.id + "/" + (this.props.parentDLName + ">" + distributionListMember.displayName)}> View</Anchor> |
                        <Anchor className="seperatorSpacing" id={distributionListMember.id} href="#" onClick={this.pinStatusUpdate}> {distributionListMember.isPinned ? "Unpin" : "Pin"}</Anchor>
                    </Flex>
                </Segment>
                )
            }
            else {
                items.push(<Segment >
                    <Flex gap="gap.small" className="actionSection">
                        <Anchor href="#" id={distributionListMember.userPrincipalName} onClick={this.oneOnOneChat}>Chat</Anchor> |
                        <Anchor className="seperatorSpacing" id={distributionListMember.id} href="#" onClick={this.pinStatusUpdate}> {distributionListMember.isPinned ? " Unpin" : " Pin"}</Anchor>
                    </Flex>
                </Segment>)
            }
        }

        let segmentRows = []; //Populate grid
        if (this.state.loader) {
            segmentRows.push(<Segment styles={{ gridColumn: 'span 5', }}>< Loader /></Segment >);
        }
        else {
            segmentRows.push(items);
        }

        const titleText = "Distribution List";

        return (
            <div className="mainComponent">
                <div className={"formContainer"}>
                    <Flex space="between">
                        <Text content={(this.props.parentDLName === "") ? titleText : titleText + ">" + this.props.parentDLName} size={"larger"} className="textstyle" /><br />

                        <Flex gap="gap.small">
                            <div className="divstyle">
                                <Dropdown
                                    fluid={true}
                                    items={pageSize}
                                    placeholder="Page Size"
                                    highlightedIndex={index}
                                    onSelectedChange={this.setPageSize}
                                    checkable
                                />
                            </div>
                            <FlexItem>
                                <Button content="Start Group Chat" disabled={!(this.checkedMembersForChat.length > 1)} primary onClick={this.groupChatWithMembers} />
                            </FlexItem>
                            <FlexItem>
                                <Input icon="search" placeholder="Search" onChange={this.search} />
                            </FlexItem>
                        </Flex>
                    </Flex>
                    <Anchor href="#" onClick={this.pageGoBack} className="textstyle" hidden={this.state.isGoBackClicked}>Go Back</Anchor>
                    <div className="formContentContainer" >
                        <Grid columns="1.5fr 1.5fr 1fr 1fr">
                            <Segment color="brand">
                                <Flex gap="gap.small">
                                    <FlexItem grow>
                                        <Checkbox key="contactName" id="contactName" label="Contact Name" onClick={this.selectAllCheckboxChanged} checked={this.state.isAllSelectChecked} />
                                    </FlexItem>
                                    <FlexItem push>
                                        <Anchor href="#" id="displayName" key="displayName" className="displayName" onClick={() => this.sortDataByColumn("displayName", this.state.nameSortIcon)}>
                                            <FontAwesomeIcon icon={this.state.nameSortIcon} />
                                        </Anchor>
                                    </FlexItem>
                                </Flex>
                            </Segment>

                            <Segment color="brand" >
                                <Flex gap="gap.small">
                                    <FlexItem grow>
                                        <Text content="Alias" />
                                    </FlexItem>

                                    <FlexItem push >
                                        <Anchor href="#" id="mail" key="mail" onClick={() => this.sortDataByColumn("mail", this.state.aliasSortIcon)}>
                                            <FontAwesomeIcon icon={this.state.aliasSortIcon} />
                                        </Anchor>
                                    </FlexItem >

                                </Flex>
                            </Segment>

                            <Segment color="brand" >
                                <Flex gap="gap.small">
                                    <FlexItem grow>
                                        <Text content="Presence Status" />
                                    </FlexItem>
                                    <Anchor href="#" id="presence" key="presence" onClick={() => this.sortDataByColumn("presence", this.state.presenceSortIcon)}>
                                        <FlexItem push >
                                            <FontAwesomeIcon icon={this.state.presenceSortIcon} />
                                        </FlexItem >
                                    </Anchor>
                                </Flex>
                            </Segment>

                            <Segment color="brand" content="Name">
                                <Flex gap="gap.small">
                                    <Text content="Actions" />
                                </Flex>
                            </Segment>

                            {segmentRows}

                        </Grid>
                    </div>
                </div>

                <div className="footerContainer">
                    <Segment className={"pagingSegment"}>
                        <Flex gap="gap.small">
                            <Pagination callbackFromParent={this.setActivePage} entitiesLength={this.state.distributionListMembers.length} activePage={this.state.activePage} numberOfContents={this.state.pageSize}></Pagination>
                        </Flex>
                    </Segment>
                </div>
            </div>

        );
    }
}
