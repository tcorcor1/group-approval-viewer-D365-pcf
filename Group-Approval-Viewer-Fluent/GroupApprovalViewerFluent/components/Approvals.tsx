import * as React from 'react';
import { useState, useEffect } from 'react';
import { IInputs } from "../generated/ManifestTypes";
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { SelectionMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { ShimmeredDetailsList,  } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IResponseBag {
    boundLookupId: string | undefined,
    approvalLookupSchemaName: string | undefined,
    requests: ComponentFramework.WebApi.Entity[],
    responses: ComponentFramework.WebApi.Entity[]
}

export interface ITableRow {
    key: string,
    name: string | undefined,
    status: JSX.Element | undefined,
    comments: string | undefined,
}

export const Approvals = (_context: ComponentFramework.Context<IInputs>) => {

    // ### SLICES
    const [ approvalState, setApprovalState ] = useState('');
    const [ isLoadingState, setIsLoadingState ] = useState(true);
    const [ tableDataState, setTableDataState ] = useState([] as ITableRow[]);
    const [ tableDataMasterState, setTableDataMasterState ] = useState([] as ITableRow[]);
    const [ columnState, setColumnState ] = useState([
        { //name
            key: 'col1',
            name: 'Name',
            fieldName: 'name',
            minWidth: 210,
            maxWidth: 350,
            isRowHeader: true,
            isResizable: true,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            data: 'string',
            isPadded: true,
        },
        { //status
            key: 'col2',
            name: 'Status',
            fieldName: 'status',
            minWidth: 100,
            maxWidth: 150,
            isRowHeader: true,
            isResizable: false,
            data: 'string',
            isPadded: true,
        },
        { //comments
            key: 'col3',
            name: 'Comments',
            fieldName: 'comments',
            minWidth: 210,
            maxWidth: 350,
            isMultiline: true,
            isRowHeader: true,
            isResizable: true,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            data: 'string',
            isPadded: true,
        },
    ] as IColumn[]);


    // ### LIFECYCLE
    useEffect(() => {
        init();
    }, []);


    // ### VARS
    // @ts-ignore
    const recordGuid:string =  _context.page.entityId;
    //@ts-ignore
    const entityType:string = _context.page.entityTypeName;
    const lookupSchemaName:string | undefined = _context.parameters.ApprovalId.attributes?.LogicalName;
    const gridTitle:string = (_context.parameters.ViewerTitle.raw!.length !== 0) 
        ? `${_context.parameters.ViewerTitle.raw as string} (${approvalState})`
        : `APPROVAL PROGRESS SUMMARY (${approvalState})`; 
    let approvalBag: IResponseBag = {
        boundLookupId: '',
        approvalLookupSchemaName: lookupSchemaName,
        requests: [],
        responses: []
    };


    // ### METHODS
    const init = async () => {  
        
        //getBoundRecordDetails
        await _context.webAPI.retrieveRecord(entityType, recordGuid, `?$select=_${approvalBag.approvalLookupSchemaName}_value`).then(
            function success(result) {
                approvalBag.boundLookupId = result[`_${lookupSchemaName}_value`];
            }, 
            function error(error) {
                setIsLoadingState(false);
                console.error('Group Approval Viewer - Retrieve Bound Lookup Error', error);
            }    
        );

        if (!approvalBag.boundLookupId) {
            setApprovalState('No Approval Assigned');
            setIsLoadingState(false);
            return;
        };

        //getApprovalRequests & setApprovalState
        await _context.webAPI.retrieveMultipleRecords("msdyn_flow_approvalrequest",`?$filter=_msdyn_flow_approvalrequest_approval_value eq ${approvalBag.boundLookupId} &$select=_ownerid_value,createdon&$expand=msdyn_flow_approvalrequest_approval($select=msdyn_flow_approval_result)`).then(
            function success(result) {
                approvalBag.requests = result.entities;
                if (approvalBag.requests.length !== 0) {
                    const approvalResult:string | null = approvalBag.requests[0].msdyn_flow_approvalrequest_approval.msdyn_flow_approval_result;
                    setApprovalState((approvalResult) ? approvalResult : 'In Progress');
                };
            },
            function error(error) {
                setIsLoadingState(false);
                console.error('Group Approval Viewer - Retrieve Approval Requests Error', error);
            }
        );

        if (approvalBag.requests.length === 0 ) {
            setIsLoadingState(false);
            setApprovalState('No Requests Found');
            return;
        };

        //getApprovalResponse
        await _context.webAPI.retrieveMultipleRecords("msdyn_flow_approvalresponse",`?$filter=_msdyn_flow_approvalresponse_approval_value eq ${approvalBag.boundLookupId} &$select=_ownerid_value,createdon,msdyn_flow_approvalresponse_response,msdyn_flow_approvalresponse_comments`).then(
            function success(result) {
                approvalBag.responses = result.entities;
            },
            function error(error) {
                setIsLoadingState(false);
                console.error('Group Approval Viewer - Retrieve Approval Responses Error', error);
            }
        );

        transformLoadTableData(); 
    };

    const transformLoadTableData = () => {
        const approvalGridArray:ITableRow[] = approvalBag.requests.map((e, index) => {
            return {
                key: index.toString(),
                name: e['_ownerid_value@OData.Community.Display.V1.FormattedValue'],
                status: getStatusIcon(e['_ownerid_value'], approvalBag.responses),
                comments: getApproverComments(e['_ownerid_value'], approvalBag.responses)
            };
        });

        setTableDataState(approvalGridArray);
        setTableDataMasterState(approvalGridArray);
        setIsLoadingState(false);
    };

    const getStatusIcon = (
        guid:string,
        respArr: ComponentFramework.WebApi.Entity[]
    ) : JSX.Element | undefined => {
        try {
            const responseItem = respArr.find((e) => e['_ownerid_value'] === guid );
            if (responseItem) {
                const responseState = responseItem!['msdyn_flow_approvalresponse_response'].toString().toLowerCase();
                switch (responseState) {
                    case 'approve':
                        return <Icon className='icon icon-approve' iconName="LikeSolid"  />;
                    case 'reject': 
                        return <Icon className='icon icon-reject' iconName="DislikeSolid"  />;
                }
            } else {
                return <Icon className='icon' iconName="Clock"  />
            };
        }
        catch(error) {
            console.error('Group Approval Viewer - Retrieve Status Icon Error', error);
        }
    };

    const getApproverComments = (
        guid:string,
        respArr: ComponentFramework.WebApi.Entity[]
    ) : string | undefined => {
        try {
            const responseItem = respArr.find((e) => e['_ownerid_value'] === guid);
            if (responseItem) {
                return responseItem!['msdyn_flow_approvalresponse_comments'].toString();
            } else {
                return '#N/A'
            };
        }
        catch(error) {
            console.error('Group Approval Viewer - Retrieve Approval Response Error', error);
        }
    };

    const onSearchHandler = (
        e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>
    ) : void => {
        const query: string = (e.target as HTMLInputElement).value;
        const itemsFiltered = tableDataState.filter(row => {
            return (
                row.name?.toLowerCase().includes(query) ||
                row.comments?.toLowerCase().includes(query)
            );
        });

        setTableDataState(query ? itemsFiltered : tableDataMasterState);
    };

    const onSortHandler = (
        ev?: React.MouseEvent<HTMLElement>, 
        column?: IColumn
    ) : void  => {
        if (!!column) {
            const newColumns: IColumn[] = [ ...columnState ];
            const currColumn: IColumn = newColumns.filter(currColumn => column.key === currColumn.key)[0];
            newColumns.forEach((newCol: IColumn) => {
                if (newCol === currColumn) {
                  currColumn.isSortedDescending = !currColumn.isSortedDescending;
                  currColumn.isSorted = true;
                } else {
                  newCol.isSorted = false;
                  newCol.isSortedDescending = true;
                }
            });
            const sortedItems = copySort(tableDataState, currColumn.fieldName!, currColumn.isSortedDescending);
            setTableDataState(sortedItems);
        }
    };

    const copySort = <T extends unknown> (
        items: T[], 
        columnKey: string, 
        isSortedDescending?: boolean
    ): T[] => {
        const key = columnKey as keyof T;
        return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
    };

    
    // ### COMPONENT
    return (
        <Fabric>
            <div className='approval-head'>
                <h2>{gridTitle}</h2>
                <TextField 
                    placeholder="Search" 
                    className='search-input'
                    onChange={onSearchHandler}
                />
            </div>
            <ShimmeredDetailsList
                setKey="items"
                items={tableDataState}
                columns={columnState}
                selectionMode={SelectionMode.none}
                enableShimmer={isLoadingState}
                ariaLabelForShimmer="Please Wait"
                ariaLabelForGrid="Approval Response Data"
                isHeaderVisible={true}
                onColumnHeaderClick={onSortHandler}
            />
        </Fabric>
    );
}; 