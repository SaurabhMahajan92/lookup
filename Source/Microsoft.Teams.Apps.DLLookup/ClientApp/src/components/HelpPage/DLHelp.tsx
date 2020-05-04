import * as React from "react";
import { Text, Accordion, Layout, Image, Flex } from "@stardust-ui/react";

export default class DLHelp extends React.Component<{}, {}> {

    private pageSize: { key: string, title: {}, content: {} }[] = [
        {
            key: 'pageSize',
            title: (<Text content="How to change the Page Size?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="pageSizeContent">
                        <Text content="Use Page Size dropdown for selecting the numbers content you want on the page." />
                        <Image circular src="/images/numbers/9.png" />
                    </Flex>
                )
        },
    ];

    private addDistributionList: { key: string, title: {}, content: {} }[] = [
        {
            key: 'addDL',
            title: (<Text content="How to add Favorite Distribution List?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="addDLContent">
                        <Text content="To add favorite distribution lists into the app, please use the Add Distribution Button on top-right." />
                        <Image circular src="/images/numbers/1.png" />
                    </Flex>
                )
        },
    ];

    private searchDistributionList: { key: string, title: {}, content: {} }[] = [
        {
            key: 'searchDL',
            title: (<Text content="How to search for your favorite Distribution List?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="searchDLContent">
                        <Text content="Search and find distribution list from your favorite distribution lists with the Search Box provided on top-right." />
                        <Image circular src="/images/numbers/5.png" />
                    </Flex>
                )
        },
    ];

    private sorting: { key: string, title: {}, content: {} }[] = [
        {
            key: 'sorting',
            title: (<Text content="How to Sort according to respective coulmns?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="sortingContent">
                        <Text content="Please click on the the icons against each heading for sort based on the given table." />
                        <Image circular src="/images/numbers/6.png" />
                    </Flex>
                )
        },
    ];

    private distributionListActions: { key: string, title: {}, content: {} }[] = [
        {
            key: 'view',
            title: (<Text content="How to view Distribution List Members?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="viewContent">
                        <Text content='Click on the link "View" to see the members of the distribution list.' />
                        <Image circular src="/images/numbers/2.png" />
                    </Flex>
                )
        },
        {
            key: 'pin',
            title: (<Text content="How to pin frequently used Distribution List?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="pinContent">
                        <Text content='Click on the link "Pin" to make sure the given distribution list comes always on top for you.' />
                        <Image circular src="/images/numbers/3.png" />
                    </Flex>
                )
        },
        {
            key: 'hide',
            title: (<Text content="How to remove a Distribution List?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="hideContent">
                        <Text content='Click on the link "Hide" to remove the given distribution list from the favourites list.' />
                        <Image circular src="/images/numbers/4.png" />
                    </Flex>
                )
        },
    ];

    public render(): JSX.Element {
        return (
            <div className="contentContainer">
                <Accordion panels={this.addDistributionList} />
                <Accordion panels={this.distributionListActions} />
                <Accordion panels={this.searchDistributionList} />
                <Accordion panels={this.sorting} />
                <Accordion panels={this.pageSize} />
                <br />
                <Layout
                    styles={{
                        width: '1000px',
                    }}
                    renderMainArea={() => <Image fluid src="/images/DistributionList.jpg" />}
                />
            </div>
        );
    }
};