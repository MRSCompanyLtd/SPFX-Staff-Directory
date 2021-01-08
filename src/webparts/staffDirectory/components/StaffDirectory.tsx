import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from "./StaffDirectory.module.scss";
import { PersonaCard } from "./PersonaCard/PersonaCard";
import { spservices } from "../../../SPServices/spservices";
import { IStaffDirectoryState } from "./IStaffDirectoryState";
import * as strings from "StaffDirectoryWebPartStrings";
import {
    Spinner, SpinnerSize, MessageBar, MessageBarType, SearchBox, Icon, Label,
    Pivot, PivotItem, PivotLinkFormat, PivotLinkSize, Dropdown, IDropdownOption
} from "office-ui-fabric-react";
import { Stack, IStackStyles, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
import { debounce } from "throttle-debounce";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ISPServices } from "../../../SPServices/ISPServices";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { spMockServices } from "../../../SPServices/spMockServices";
import { IStaffDirectoryProps } from './IStaffDirectoryProps';
import Paging from './Pagination/Paging';

const slice: any = require('lodash/slice');
const filter: any = require('lodash/filter');
const wrapStackTokens: IStackTokens = { childrenGap: 30 };

const StaffDirectory: React.FC<IStaffDirectoryProps> = (props) => {
    let _services: ISPServices = null;
    if (Environment.type === EnvironmentType.Local) {
        _services = new spMockServices();
    } else {
        _services = new spservices(props.context);
    }
    const [az, setaz] = useState<string[]>([]);
    const [alphaKey, setalphaKey] = useState<string>('A');
    const [state, setstate] = useState<IStaffDirectoryState>({
        users: [],
        isLoading: true,
        errorMessage: "",
        hasError: false,
        indexSelectedKey: "A",
        searchString: "LastName",
        searchText: "",
        department: ""
    });
    const departmentOptions: IDropdownOption[] = props.departments ? props.departments.map(d => ({key: d.departmentKey, text: d.departmentName})) : [];
    const color = props.context.microsoftTeams ? "white" : "";
    // Paging
    const [pagedItems, setPagedItems] = useState<any[]>([]);
    const [pageSize, setPageSize] = useState<number>(props.pageSize ? props.pageSize : 10);
    const [currentPage, setCurrentPage] = useState<number>(1);

    const _onPageUpdate = async (pageno?: number) => {
        var currentPge = (pageno) ? pageno : currentPage;
        var startItem = ((currentPge - 1) * pageSize);
        var endItem = currentPge * pageSize;
        let filItems = slice(state.users, startItem, endItem);
        setCurrentPage(currentPge);
        setPagedItems(filItems);
    };

    const directoryGrid =
        pagedItems && pagedItems.length > 0
            ? pagedItems.map((user: any) => {
                return (
                    <PersonaCard
                        context={props.context}
                        profileProperties={{
                            DisplayName: user.PreferredName,
                            Title: user.JobTitle,
                            PictureUrl: user.PictureURL,
                            Email: user.WorkEmail,
                            Department: user.Department,
                            WorkPhone: user.WorkPhone,
                            Location: user.OfficeNumber
                                ? user.OfficeNumber
                                : user.BaseOfficeLocation
                        }}
                    />
                );
            })
            : [];
    const _loadAlphabets = () => {
        let alphabets: string[] = [];
        for (let i = 65; i < 91; i++) {
            alphabets.push(
                String.fromCharCode(i)
            );
        }
        setaz(alphabets);
    };

    const _alphabetChange = async (item?: PivotItem, ev?: React.MouseEvent<HTMLElement>) => {
        setstate({ ...state, searchText: "", indexSelectedKey: item.props.itemKey, isLoading: true });
        setalphaKey(item.props.itemKey);
        setCurrentPage(1);
    };
    const _searchByAlphabets = async (initialSearch: boolean) => {
        setstate({ ...state, isLoading: true, searchText: '' });
        let users = null;
        if (initialSearch) {
            if (props.searchFirstName)
                users = await _services.searchUsersNew('', `FirstName:a*`, false, props.query);
            else users = await _services.searchUsersNew('a', '', true, props.query);
        } else {
            if (props.searchFirstName)
                users = await _services.searchUsersNew('', `FirstName:${alphaKey}*`, false, props.query);
            else users = await _services.searchUsersNew(`${alphaKey}`, '', true, props.query);
        }
        setstate({
            ...state,
            searchText: '',
            indexSelectedKey: initialSearch ? 'A' : state.indexSelectedKey,
            users:
                users && users.PrimarySearchResults
                    ? users.PrimarySearchResults
                    : null,
            isLoading: false,
            errorMessage: "",
            hasError: false
        });
    };

    let _searchUsers = async (searchText: string) => {
        try {
            setstate({ ...state, searchText: searchText, isLoading: true });
            if (searchText.length > 0) {
                let searchProps: string[] = props.searchProps && props.searchProps.length > 0 ?
                    props.searchProps.split(',') : ['FirstName', 'LastName', 'WorkEmail', 'Department'];
                let qryText: string = '';
                let finalSearchText: string = searchText ? searchText.replace(/ /g, '+') : searchText;
                if (props.clearTextSearchProps) {
                    let tmpCTProps: string[] = props.clearTextSearchProps.indexOf(',') >= 0 ? props.clearTextSearchProps.split(',') : [props.clearTextSearchProps];
                    if (tmpCTProps.length > 0) {
                        searchProps.map((srchprop, index) => {
                            let ctPresent: any[] = filter(tmpCTProps, (o) => { return o.toLowerCase() == srchprop.toLowerCase(); });
                            if (ctPresent.length > 0) {
                                if(index == searchProps.length - 1) {
                                    qryText += `${srchprop}:${searchText}*`;
                                } else qryText += `${srchprop}:${searchText}* OR `;
                            } else {
                                if(index == searchProps.length - 1) {
                                    qryText += `${srchprop}:${finalSearchText}*`;
                                } else qryText += `${srchprop}:${finalSearchText}* OR `;                                
                            }
                        });
                    } else {
                        searchProps.map((srchprop, index) => {
                            if (index == searchProps.length - 1)
                                qryText += `${srchprop}:${finalSearchText}*`;
                            else qryText += `${srchprop}:${finalSearchText}* OR `;
                        });
                    }
                } else {
                    searchProps.map((srchprop, index) => {
                        if (index == searchProps.length - 1)
                            qryText += `${srchprop}:${finalSearchText}*`;
                        else qryText += `${srchprop}:${finalSearchText}* OR `;
                    });
                }
                const users = await _services.searchUsersNew('', qryText, false, props.query);
                setstate({
                    ...state,
                    searchText: searchText,
                    indexSelectedKey: '0',
                    users:
                        users && users.PrimarySearchResults
                            ? users.PrimarySearchResults
                            : null,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });
                setalphaKey('0');
            } else {
                setstate({ ...state, searchText: '' });
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _searchBoxChanged = (newvalue: string): void => {
        setCurrentPage(1);
        _searchUsers(newvalue);
        console.log(state);
    };
    _searchUsers = debounce(500, _searchUsers);

    const _searchDepartment = async (newvalue: string) => {
        await _searchUsers(newvalue);
    };

    useEffect(() => {
        setPageSize(props.pageSize);
        if (state.users) _onPageUpdate();
    }, [state.users, props.pageSize]);

    useEffect(() => {
        if (alphaKey.length > 0 && alphaKey != "0") _searchByAlphabets(false);
    }, [alphaKey]);

    useEffect(() => {
        _loadAlphabets();
        _searchByAlphabets(true);
    }, [props]);

    return (
        <div className={styles.directory}>
            <WebPartTitle displayMode={props.displayMode} title={props.title}
                updateProperty={props.updateProperty} />
            <div className={styles.searchBox}>
                <SearchBox placeholder={strings.SearchPlaceHolder} className={styles.searchTextBox}
                    onSearch={_searchUsers}
                    value={state.searchText}
                    onChange={_searchBoxChanged} />
                <div>
                {
                props.departmentFilter &&
                <div className={styles.dropDownSortBy}>
                    <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
                        <Dropdown
                            key="departmentDropdown"
                            placeholder="Select department"
                            label="Filter by department"
                            options={departmentOptions}
                            onChange={(ev: any, value: IDropdownOption) => {
                                _searchDepartment(value.key.toString());
                            }}
                            styles={{ dropdown: { width: 200, textAlign: 'left' } }}
                        />
                    </Stack>
                </div>
                }
                    <Pivot className={styles.alphabets} linkFormat={PivotLinkFormat.tabs}
                        selectedKey={state.indexSelectedKey} onLinkClick={_alphabetChange}
                        linkSize={PivotLinkSize.normal} >
                        {az.map((index: string) => {
                            return (
                                <PivotItem headerText={index} itemKey={index} key={index} />
                            );
                        })}
                    </Pivot>
                </div>
            </div>
            {state.isLoading ? (
                <div style={{ marginTop: '10px' }}>
                    <Spinner size={SpinnerSize.large} label={strings.LoadingText} />
                </div>
            ) : (
                    <>
                        {state.hasError ? (
                            <div style={{ marginTop: '10px' }}>
                                <MessageBar messageBarType={MessageBarType.error}>
                                    {state.errorMessage}
                                </MessageBar>
                            </div>
                        ) : (
                                <>
                                    {!pagedItems || pagedItems.length == 0 ? (
                                        <div className={styles.noUsers}>
                                            <Icon
                                                iconName={"ProfileSearch"}
                                                style={{ fontSize: "54px", color: color }}
                                            />
                                            <Label>
                                                <span style={{ marginLeft: 5, fontSize: "26px", color: color }}>
                                                    {strings.DirectoryMessage}
                                                </span>
                                            </Label>
                                        </div>
                                    ) : (
                                            <>
                                                <div style={{ width: '100%', display: 'inline-block' }}>
                                                    <Paging
                                                        totalItems={state.users.length}
                                                        itemsCountPerPage={pageSize}
                                                        onPageUpdate={_onPageUpdate}
                                                        currentPage={currentPage} />
                                                </div>
                                                <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
                                                    <div>{directoryGrid}</div>
                                                </Stack>
                                                <div style={{ width: '100%', display: 'inline-block' }}>
                                                    <Paging
                                                        totalItems={state.users.length}
                                                        itemsCountPerPage={pageSize}
                                                        onPageUpdate={_onPageUpdate}
                                                        currentPage={currentPage} />
                                                </div>
                                            </>
                                        )}
                                </>
                            )}
                    </>
                )}
        </div>
    );
};

export default StaffDirectory;