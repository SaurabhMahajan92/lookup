import * as React from "react";
import { Text, Accordion, Layout, Image, Flex } from "@stardust-ui/react";

export default class DLMHelp extends React.Component<{}, {}> {
    private pageSize: { key: string, title: {}, content: {} }[] = [
        {
            key: 'pageSize',
            title: (<Text content="How to change the Page Size?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="pageSizeContent">
                        <Text content='Use Page Size dropdown for selecting the numbers content you want on the page.' />
                        <Image circular src="/images/numbers/9.png" />
                    </Flex>
                )
        },
    ];

    private searchDL: { key: string, title: {}, content: {} }[] = [
        {
            key: 'searchDL',
            title: (<Text content="How to search for your favorite Contact?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="searchDLContent">
                        <Text content='Search and find contact from distribution list members with the Search Box provided on top-right.' />
                        <Image circular src="/images/numbers/6.png" />
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
                        <Text content='Please click on the the icons against each heading for sort based on the given table.' />
                        <Image circular src="/images/numbers/8.png" />
                    </Flex>
                )
        },
    ];

    private dlMemberActions: { key: string, title: {}, content: {} }[] = [
        {
            key: 'view',
            title: (<Text content="How to check if the member of the distribution List is another Distribution list?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="viewContent">
                        <Text content='If the distribution list member is another distribution list, you can see View option' />
                        <Image circular src="/images/numbers/7.png" />
                        <Text content='which will take you to the members list of that distribution list.' />
                    </Flex>
                )
        },
        {
            key: 'chat',
            title: (<Text content="How to Chat with a member of the Distribution List?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="chatContent">
                        <Text content='If the member is a contact, you will see Chat option.' />
                        <Image circular src="/images/numbers/1.png" />

                    </Flex>
                )
        },
        {
            key: 'pin',
            title: (<Text content="How to pin/unpin frequently contacted Distribution List Member?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="pinContent">
                        <Text content='Click on the link "Pin/Unpin" so that your favorite contacts will appear on top of the list next time you visit.' />
                        <Image circular src="/images/numbers/2.png" />
                    </Flex>
                )
        },
    ];

    private groupChat: { key: string, title: {}, content: {} }[] = [
        {
            key: 'groupChat',
            title: (<Text content="How to start a group chat with multiple members?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="groupChatContent">
                        <Text content='Select the respective members with whom you want to start a group chat by clicking on the checkbox near their name' />
                        <Image circular src="/images/numbers/5.png" />
                        <Text content='  and click on the Start Group Chat button' />
                        <Image circular src="/images/numbers/3.png" />
                    </Flex>
                )
        },
        {
            key: 'groupChatAll',
            title: (<Text content="How to start a group chat with all members?" className="fontstyle" />),
            content:
                (
                    <Flex gap="gap.smaller" key="groupChatAllContent">
                        <Text content='You can select all the members in a distribution list by clicking on checkbox near the Contact Name column' />
                        <Image circular src="/images/numbers/4.png" />
                        <Text content='  and click on the Start Group Chat button' />
                        <Image circular src="/images/numbers/3.png" />
                    </Flex>
                )
        },
    ];

    public render(): JSX.Element {
        return (
            <div className="contentContainer">
                <Accordion panels={this.dlMemberActions} />
                <Accordion panels={this.groupChat} />
                <Accordion panels={this.searchDL} />
                <Accordion panels={this.sorting} />
                <Accordion panels={this.pageSize} />
                <br />
                <Layout
                    styles={{
                        width: '1000px',
                    }}
                    renderMainArea={() => <Image fluid src="/images/DLMembers.jpg" />}
                />
            </div>
        );
    }
};