import axios from './axiosJWTDecorator';
import { getBaseUrl } from '../configVariables';
import { AxiosResponse } from "axios";
import { IADDistributionList } from "./../components/AddDistributionList/AddDistributionList"
import { IDistributionListMember, IUserPageSizeChoice, IPresenceData } from "./../components/DistributionListMembers/DistributionListMembers"
import { IDistributionList } from "./../components/DistributionLists/DistributionLists"

let baseAxiosUrl = getBaseUrl() + '/api';

export const getFavoriteDistributionLists = async (): Promise<AxiosResponse<IDistributionList[]>> => {
    let url = baseAxiosUrl + "/distributionlists";
    return await axios.get(url);
}

export const getADDistributionLists = async (query: string): Promise<AxiosResponse<IADDistributionList[]>> => {
    let url = baseAxiosUrl + "/distributionlists/getDistributionList?query=" + query;
    return await axios.get(url);
}

export const createFavoriteDistributionList = async (payload: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/distributionlists";
    return await axios.post(url, payload);
}

export const updateFavoriteDistributionList = async (payload: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/distributionlists";
    return await axios.put(url, payload);
}

export const deleteFavoriteDistributionList = async (payload: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/distributionlists";
    return await axios.delete(url, payload);
}

export const getDistributionListsMembers = async (groupID?: string): Promise<AxiosResponse<IDistributionListMember[]>> => {
    let url = baseAxiosUrl + "/distributionlistmembers?groupID=" + groupID;
    return await axios.get(url);
}

export const pinStatusUpdate = async (pinnedUser: string, status: boolean, distributionListID: string): Promise<AxiosResponse<void>> => {
    var payload = {
        "pinnedUserId": pinnedUser,
        "distributionListID": distributionListID
    }
    if (status) {
        let url = baseAxiosUrl + "/distributionlistmembers";
        return await axios.post(url, payload);
    }
    else {
        let url = baseAxiosUrl + "/distributionlistmembers";
        return await axios.delete(url, payload);
    }
}

export const getDistributionListMembersOnlineCount = async (groupId?: string): Promise<AxiosResponse<string>> => {
    let url = baseAxiosUrl + "/presence/GetDistributionListMembersOnlineCount?groupId=" + groupId;
    return await axios.get(url);
}

export const getUserPresence = async (payload: {}): Promise<AxiosResponse<IPresenceData[]>> => {
    let url = baseAxiosUrl + "/presence/getUserPresence";
    return await axios.post(url, payload);
}

export const getUserPageSizeChoice = async (): Promise<AxiosResponse<IUserPageSizeChoice>> => {
    let url = baseAxiosUrl + "/UserPageSize";
    return await axios.get(url);
}

export const createUserPageSizeChoice = async (payload: {}): Promise<AxiosResponse<void>> => {
    let url = baseAxiosUrl + "/UserPageSize";
    return await axios.post(url, payload);
}

export const getAuthenticationMetadata = async (windowLocationOriginDomain: string, loginHint: string): Promise<AxiosResponse<string>> => {
    let url = `${baseAxiosUrl}/authenticationMetadata/GetAuthenticationUrlWithConfiguration?windowLocationOriginDomain=${windowLocationOriginDomain}&loginhint=${loginHint}`;
    return await axios.get(url, undefined, false);
}

export const getClientId = async (): Promise<AxiosResponse<string>> => {
    let url = baseAxiosUrl + "/authenticationMetadata/getClientId";
    return await axios.get(url);
}