import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp-admin";
import "@pnp/sp/search";
import { ISearchQuery, SearchResults, SearchQueryBuilder, ISearchResult, Search, ISearchBuilder, SortDirection } from "@pnp/sp/search";
import * as React from "react";
import "@pnp/graph/users";
import { getSP } from "./pnpjsConfig";
import { GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph";
import { ISitePermissionsProps } from "./ISitePermissionsProps";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/profiles";
import "@pnp/graph/groups";
import "@pnp/graph/members";
import "@pnp/sp/folders/web";
import "@pnp/graph/sites";
import "@pnp/sp/lists";
import "@pnp/sp/security";
import { SharingLinkKind, IShareLinkResponse } from "@pnp/sp/sharing";
import "@pnp/sp/sharing";
import "@pnp/sp/folders";
import "@pnp/sp/site-groups"
import { ISharingInformation } from "@pnp/sp/sharing";
import { PermissionKind } from "@pnp/sp/security";



export interface IPermissionsObject {
    visibility: string;
    uniquePermissions: boolean;
    externalUsers: boolean;
    externalEnabled: boolean;
    members: Array<IMembersObject>;
    owners: Array<IMembersObject>;
    guests: Array<IMembersObject>;
}

export interface IMembersObject {
    name: string;
    mail: string;
    userName: string;
    id: string;
    edit: boolean;
    view: boolean;
    fromGraph: boolean;
}

export interface IInheritanceList {
    listId: string;
    itemId: number;
}




export default class PermissionsSearch extends React.Component<ISitePermissionsProps>{

    public returnPermissions: IPermissionsObject;

    private _sp: SPFI;
    private graph: GraphFI;
    public context: WebPartContext;

    private inheritList: Array<IInheritanceList>;
    private inheritItemList: Array<IInheritanceList>;
    private guests: any;
    private members: any;
    private owners: any;
    private spMembers: any;
    private spOwners: any
    private allGroups: any;


    constructor(context: WebPartContext, props) {
        super(props);
        this.context = context;
        this.graph = graphfi().using(graphSPFx(context));
        this.inheritItemList = new Array();
        this.inheritList = new Array();
        this.returnPermissions = { externalEnabled: undefined, visibility: undefined, uniquePermissions: false, externalUsers: false, members: new Array(), owners: new Array(), guests: new Array() };
    }

    public async createPermissionsObject() {
        this._sp = getSP();


        await this.getGroups();
        await this.mapMembers();
        await this.getRoleInheritance();
        await this.checkPermissions();

        console.log(this.returnPermissions)

    }

    private async getGroups() {
        let guestGroupId;
        let groupID = this.context.pageContext.site.group.id._guid;
        let data = await Promise.all([this.graph.groups.getById(groupID)(), this.graph.groups.getById(groupID).members(), this.graph.groups.getById(groupID).owners(), this._sp.web.siteGroups(), this._sp.web.associatedMemberGroup.users(), this._sp.web.associatedOwnerGroup.users(),])
        this.allGroups = data[0];
        this.returnPermissions.externalEnabled = this.context.pageContext.legacyPageContext.guestsEnabled;
        this.returnPermissions.visibility = this.context.pageContext.legacyPageContext.groupType;
        this.members = data[1];
        this.owners = data[2];
        data[3].map(async (group, index) => {
            if (group.LoginName.search("Visitors") >= 0) {
                guestGroupId = group.Id;
            }
        });

        if (guestGroupId != undefined) {
            this.guests = await this._sp.web.siteGroups.getById(guestGroupId).users();
        }

        this.spMembers = data[4];
        this.spOwners = data[5];
    }

    private async mapMembers() {
        this.returnPermissions.guests = [];
        this.returnPermissions.members = [];
        this.returnPermissions.owners = [];

        let removeMail = this.context.pageContext.web.description;
        let temp = this.context.pageContext;



        if (this.guests) {
            this.guests.map((guest, index) => {
                if (!this.members.some(member => member.mail === guest.Email)) {
                    let pushMember: IMembersObject = {
                        name: guest.Title,
                        mail: guest.Email,
                        userName: guest.UserPrincipalName,
                        id: guest.LoginName,
                        fromGraph: false,
                        edit: undefined,
                        view: undefined,
                    };
                    this.returnPermissions.guests.push(pushMember)
                }
            });
        }

        if (this.spMembers) {
            this.spMembers.map((spMember, index) => {
                if (!this.members.some(member => member.mail === spMember.Email) && spMember.Email.search(removeMail) == -1) {
                    let pushMember: IMembersObject = {
                        name: spMember.Title,
                        mail: spMember.Email,
                        userName: spMember.UserPrincipalName,
                        id: spMember.LoginName,
                        edit: undefined,
                        view: undefined,
                        fromGraph: false,
                    };
                    this.returnPermissions.members.push(pushMember)
                }
            });
        }

        if (this.spOwners) {
            this.spOwners.map((spOwner, index) => {
                if (!this.owners.some(member => member.mail === spOwner.Email) && spOwner.Email.search(removeMail) == -1 && spOwner.Email != "") {
                    let pushMember: IMembersObject = {
                        name: spOwner.Title,
                        mail: spOwner.Email,
                        userName: spOwner.UserPrincipalName,
                        id: spOwner.LoginName,
                        edit: undefined,
                        view: undefined,
                        fromGraph: false,
                    };
                    this.returnPermissions.owners.push(pushMember)
                }
            });
        }

        if (this.members) {
            this.members.map((member, index) => {
                if (!this.owners.some(owner => owner.id === member.id)) {
                    let pushMember: IMembersObject = {
                        name: member.displayName,
                        mail: member.mail,
                        userName: member.userPrincipalName,
                        id: member.id,
                        fromGraph: true,
                        edit: undefined,
                        view: undefined,
                    };
                    this.returnPermissions.members.push(pushMember)

                } else if (this.owners.some(owner => owner.id === member.id)) {
                    let pushMember: IMembersObject = {
                        name: member.displayName,
                        mail: member.mail,
                        userName: member.userPrincipalName,
                        id: member.id,
                        fromGraph: true,
                        edit: undefined,
                        view: undefined,
                    };
                    this.returnPermissions.owners.push(pushMember)
                }
            });
        }


        this.returnPermissions.guests.map((guest, index) => {
            if (guest.userName.toUpperCase().search("#EXT#") >= 0) {
                this.returnPermissions.externalUsers = true;
                return;
            }
        });


        if (!this.returnPermissions.externalUsers) {

            this.returnPermissions.members.map((guest, index) => {
                if (guest.userName.toUpperCase().search("#EXT#") >= 0) {
                    this.returnPermissions.externalUsers = true;
                    return;
                }
            });

        }


        else if (!this.returnPermissions.externalUsers) {

            this.returnPermissions.owners.map((guest, index) => {
                if (guest.userName.toUpperCase().search("#EXT#") >= 0) {
                    this.returnPermissions.externalUsers = true;
                    return;
                }
            });

        }
    }


    //Reset unique Permissions
    public async resetRoleInheritance() {

        await this.inheritList.map(async (list, index) => {
            await this._sp.web.lists.getById(list.listId).breakRoleInheritance();
            await this._sp.web.lists.getById(list.listId).resetRoleInheritance();

        });
        await this.inheritItemList.map(async (list, index) => {
            await this._sp.web.lists.getById(list.listId).items.getById(list.itemId).breakRoleInheritance();
            await this._sp.web.lists.getById(list.listId).items.getById(list.itemId).resetRoleInheritance();
        });

    }

    //get unique permissions
    private async getRoleInheritance() {

        let rolesLists: any = await this._sp.web.lists.select("Id, HasUniqueRoleAssignments")();
        await rolesLists.map(async (list, index) => {
            if (list.HasUniqueRoleAssignments) {
                let pushObj: IInheritanceList = {
                    itemId: undefined,
                    listId: list.Id,
                };
                if (!this.returnPermissions.uniquePermissions) {
                    this.returnPermissions.uniquePermissions = true;
                }

                this.inheritList.push(pushObj);
            }
        });

        let lists: any = await this._sp.web.lists();
        await lists.map(async (list, index) => {
            if (list.Hidden == false) {
                const temp = await this._sp.web.lists.getByTitle(list.Title).items.select("Id, HasUniqueRoleAssignments")();
                temp.map(async (item, index) => {
                    if (item.HasUniqueRoleAssignments) {
                        let pushObj: IInheritanceList = {
                            itemId: item.Id,
                            listId: list.Id,
                        }
                        if (!this.returnPermissions.uniquePermissions) {
                            this.returnPermissions.uniquePermissions = true;
                        }
                        this.inheritItemList.push(pushObj);
                    }
                });
            }
        });
    }

    private async checkPermissions() {
        this.returnPermissions.members.map(async (member, index) => {
            let loginName = "i:0#.f|membership|" + member.userName;
            let testSearch = await this._sp.web.getUserEffectivePermissions(loginName);
            member.edit = this._sp.web.hasPermissions(testSearch, PermissionKind.EditListItems);
            member.view = this._sp.web.hasPermissions(testSearch, PermissionKind.ViewPages);
        });

        this.returnPermissions.guests.map(async (member, index) => {
            let loginName = "i:0#.f|membership|" + member.userName;
            let testSearch = await this._sp.web.getUserEffectivePermissions(loginName);
            member.edit = this._sp.web.hasPermissions(testSearch, PermissionKind.EditListItems);
            member.view = this._sp.web.hasPermissions(testSearch, PermissionKind.ViewPages);
        });

        this.returnPermissions.owners.map(async (member, index) => {
            let loginName = "i:0#.f|membership|" + member.userName;
            let testSearch = await this._sp.web.getUserEffectivePermissions(loginName);
            member.edit = this._sp.web.hasPermissions(testSearch, PermissionKind.EditListItems);
            member.view = this._sp.web.hasPermissions(testSearch, PermissionKind.ViewPages);
        });
    }


}