import * as React from 'react';
import { useState, useEffect } from 'react';
import { IInputs } from "../generated/ManifestTypes";
import Table from "./Table";
import ThumbUpIcon from '@material-ui/icons/ThumbUp';
import ThumbDownIcon from '@material-ui/icons/ThumbDown';
import AccessTime from '@material-ui/icons/AccessTime';

export interface IResponseBag {
    boundLookupId: string | undefined,
    approvalLookupSchemaName: string | undefined,
    requests: ComponentFramework.WebApi.Entity[],
    responses: ComponentFramework.WebApi.Entity[]
}

const Approvals = (_context: ComponentFramework.Context<IInputs>) => {

    // ### SLICES
    const [ tableDataState, setTableDataState ] = useState({ tableData:[] });
    const [ isLoading, setIsLoadingState ] = useState(true);
    const [ approvalState, setApprovalState ] = useState('');

    // ### LIFECYCLE
    useEffect(() => {
        buildTable();
    }, []);

    // ### VARS
    //@ts-ignore
    const recordGuid:string =  _context.page.entityId;
    //@ts-ignore
    const entityType:string = _context.page.entityTypeName;
    const lookupSchemaName:string | undefined = _context.parameters.ApprovalId.attributes?.LogicalName;
    const gridTitle:string = (_context.parameters.ViewerTitle.raw!.length !== 0) 
        ? `${_context.parameters.ViewerTitle.raw as string} (${approvalState})`
        : `Approval Progress Summary (${approvalState})`; 
    let approvalBag: IResponseBag = {
        boundLookupId: '',
        approvalLookupSchemaName: lookupSchemaName,
        requests: [],
        responses: []
    };

    // ### METHODS
    const buildTable = async () => {
        
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

        //getApprovalResponses
        await _context.webAPI.retrieveMultipleRecords("msdyn_flow_approvalresponse",`?$filter=_msdyn_flow_approvalresponse_approval_value eq ${approvalBag.boundLookupId} &$select=_ownerid_value,createdon,msdyn_flow_approvalresponse_response,msdyn_flow_approvalresponse_comments`).then(
            function success(result) {
                approvalBag.responses = result.entities;
            },
            function error(error) {
                setIsLoadingState(false);
                console.error('Group Approval Viewer - Retrieve Approval Responses Error', error);
            }
        );

        formatTableData();
    };

    const formatTableData = () => {
        const approvalGridArray:any = approvalBag.requests.map((e) => {
            return {
                name: e['_ownerid_value@OData.Community.Display.V1.FormattedValue'],
                status: getStatusIcon(e['_ownerid_value'], approvalBag.responses),
                comments: getApproverComments(e['_ownerid_value'], approvalBag.responses)
            };
        });
        setTableDataState({ tableData: approvalGridArray });
        setIsLoadingState(false); 
    };

    const getStatusIcon = (
        guid:string,
        respArr: ComponentFramework.WebApi.Entity[]
    ) => {
        try {
            const responseItem = respArr.find((e) => e['_ownerid_value'] === guid );
            if (responseItem) {
                const responseState = responseItem!['msdyn_flow_approvalresponse_response'].toString().toLowerCase();
                switch (responseState) {
                    case 'approve':
                        return <ThumbUpIcon className={'is-approved'} />;
                    case 'reject': 
                        return <ThumbDownIcon className={'is-rejected'} />;
                }
            } else {
                return <AccessTime />
            };
        }
        catch(error) {
            console.error('Group Approval Viewer - Retrieve Status Icon Error', error);
        }
    };

    const getApproverComments = (
        guid:string,
        respArr: ComponentFramework.WebApi.Entity[]
    ) => {
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

    // ### component
    return (
        <div className='approvals-wrapper'>
            <div className='approvals-content'>
                <Table 
                    tableData={tableDataState.tableData}
                    isLoading={isLoading}
                    title={gridTitle}
                />
            </div>
        </div>
    );
}; 

export default Approvals;