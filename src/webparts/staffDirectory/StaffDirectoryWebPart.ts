import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneToggle,
    PropertyPaneSlider
} from "@microsoft/sp-property-pane";
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

import * as strings from "StaffDirectoryWebPartStrings";
import StaffDirectory from "./components/StaffDirectory";
import { IStaffDirectoryProps } from "./components/IStaffDirectoryProps";

export interface IStaffDirectoryWebPartProps {
    title: string;
    searchFirstName: boolean;
    searchProps: string;
    clearTextSearchProps: string;
    pageSize: number;
    query: string;
    departmentFilter: boolean;
    departments: any[];
}

export interface IPropertyControlsDepartments {
    departments: any[];
}

export default class DirectoryWebPart extends BaseClientSideWebPart<
    IStaffDirectoryWebPartProps
    > {
    public render(): void {
        const element: React.ReactElement<IStaffDirectoryProps> = React.createElement(
            StaffDirectory,
            {
                title: this.properties.title,
                context: this.context,
                searchFirstName: this.properties.searchFirstName,
                displayMode: this.displayMode,
                updateProperty: (value: string) => {
                    this.properties.title = value;
                },
                searchProps: this.properties.searchProps,
                clearTextSearchProps: this.properties.clearTextSearchProps,
                pageSize: this.properties.pageSize,
                query: this.properties.query,
                departmentFilter: this.properties.departmentFilter,
                departments: this.properties.departments
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse("1.0");
    }

    protected get disableReactivePropertyChanges(): boolean {
        return true;
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField("title", {
                                    label: strings.TitleFieldLabel
                                }),
                                PropertyPaneToggle("searchFirstName", {
                                    checked: false,
                                    label: "Search on First Name ?"
                                }),
                                PropertyPaneTextField('searchProps', {
                                    label: strings.SearchPropsLabel,
                                    description: strings.SearchPropsDesc,
                                    value: this.properties.searchProps,
                                    multiline: false,
                                    resizable: false
                                }),
                                PropertyPaneTextField('clearTextSearchProps', {
                                    label: strings.ClearTextSearchPropsLabel,
                                    description: strings.ClearTextSearchPropsDesc,
                                    value: this.properties.clearTextSearchProps,
                                    multiline: false,
                                    resizable: false
                                }),
                                PropertyPaneSlider('pageSize', {
                                    label: 'Results per page',
                                    showValue: true,
                                    max: 20,
                                    min: 2,
                                    step: 2,
                                    value: this.properties.pageSize
                                }),
                                PropertyPaneTextField('query', {
                                    label: "Appended filter query",
                                    description: "Add a filter query to every search",
                                    value: this.properties.query,
                                    multiline: false,
                                    resizable: false
                                }),
                                PropertyPaneToggle('departmentFilter', {
                                    label: "Add department filter",
                                    onText: "Show filter",
                                    offText: "No filter",
                                    checked: this.properties.departmentFilter
                                }),
                                PropertyFieldCollectionData('departments', {
                                    key: "departments",
                                    label: 'Departments',
                                    panelHeader: 'Department list',
                                    manageBtnLabel: 'Manage departments',
                                    value: this.properties.departments,
                                    fields: [
                                        {
                                            id: "departmentKey",
                                            title: "Department key (as defined in AD)",
                                            type: CustomCollectionFieldType.string,
                                            required: true
                                        },
                                        {
                                            id: "departmentName",
                                            title: "Department Name",
                                            type: CustomCollectionFieldType.string,
                                            required: true
                                        }
                                    ]
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
