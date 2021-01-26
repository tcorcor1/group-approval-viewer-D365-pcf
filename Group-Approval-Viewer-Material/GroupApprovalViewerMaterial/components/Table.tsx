import * as React from 'react';
import { forwardRef } from 'react';
import MaterialTable from 'material-table';
import FirstPage from '@material-ui/icons/FirstPage';
import LastPage from '@material-ui/icons/LastPage';
import SaveAlt from '@material-ui/icons/SaveAlt';
import PreviousPage from '@material-ui/icons/ChevronLeft';
import NextPage from '@material-ui/icons/ChevronRight';
import ArrowDownward from '@material-ui/icons/ArrowDownward';
import Clear from '@material-ui/icons/Clear';
import Search from '@material-ui/icons/Search';

export interface ITableProps {
  tableData: any[],
  isLoading: boolean,
  title: string
}

const Table = (props: ITableProps) => { 
  return (
      <MaterialTable 
        title={props.title} 
        data={props.tableData}
        isLoading={props.isLoading}
        columns={[
          {
            title: "Name",
            field: "name",
            type: "string",
            cellStyle: {
              width: '30%',
            },
          },
          {
            title: "Status",
            field: "status",
            cellStyle: {
              width: '20%',
            }
          },
          {
            title: "Comments",
            field: "comments",
            type: "string",
            cellStyle: {
              width: '50%',
            }
          }
        ]}
        icons={{
          FirstPage: forwardRef((props, ref) => <FirstPage {...props} ref={ref} />),
          LastPage: forwardRef((props, ref) => <LastPage {...props} ref={ref} />),
          PreviousPage: forwardRef((props, ref) => <PreviousPage {...props} ref={ref} />),
          NextPage: forwardRef((props, ref) => <NextPage {...props} ref={ref} />),
          Export: forwardRef((props, ref) => <SaveAlt {...props} ref={ref} />),
          Search: forwardRef((props, ref) => <Search {...props} ref={ref} />),
          ResetSearch: forwardRef((props, ref) => <Clear {...props} ref={ref} />),
          SortArrow: forwardRef((props, ref) => <ArrowDownward {...props} ref={ref} />)
        }}
        options={{
          search: true,
          showTitle: true,
          filtering: false,
          sorting: false,
          exportButton: false,
        }}
      />
  );
}

export default Table;