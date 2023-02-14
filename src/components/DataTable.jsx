import { IconButton, TextField } from '@mui/material'
import { DataGrid, GridToolbar } from '@mui/x-data-grid'
import React, { useEffect, useState } from 'react'
import ClearIcon from '@mui/icons-material/Clear';
import SearchIcon from '@mui/icons-material/Search';

function QuickSearchToolbar(props) {

  return (
    <div
      sx={{
        p: 1,
        display: 'flex',
        alignItems: 'flex-start',
        flexWrap: 'wrap',
        justifyContent: 'space-between',
      }}
    >
      <div>
        <GridToolbar />
      </div>
      <TextField
        variant="standard"
        value={props.value}
        onChange={props.onChange}
        placeholder="Search…"
        sx={{
          width: "100%",
          mr: 1,
          borderBottom: `1px solid`,
          borderColor: 'primary.main'
        }}
        InputProps={{
          startAdornment: <SearchIcon fontSize="small" />,
          endAdornment: (
            <IconButton
              title="Clear"
              aria-label="Clear"
              size="small"
              style={{ visibility: props.value ? 'visible' : 'hidden' }}
              onClick={props.clearSearch}
            >
              <ClearIcon fontSize="small" />
            </IconButton>
          ),
        }}
      />
    </div>
  );
}

export const DataTable = (props) => {
  let {originalRows, columns} = props
  function escapeRegExp(value) {
    return value.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
  }

  const [rows, setRows] = useState(originalRows)
  const [searchText, setSearchText] = useState("")

  useEffect(() => {
    setRows(originalRows)
    setSearchText("")
  }, [originalRows])


  const requestSearch = (searchValue) => {
    setSearchText(searchValue);
    const searchRegex = new RegExp(escapeRegExp(searchValue), 'i');
    const filteredRows = originalRows.filter((row) => {
      let rowWithoutId = { ...row }
      delete rowWithoutId.id
      return Object.keys(rowWithoutId).some((field) => {
        return searchRegex.test((row[field]) ? row[field].toString() : "");
      });
    });
    setRows(filteredRows);
  };

  return (
    <div style={{ height: 600, width: '100%' }}>
      <DataGrid
        rows={rows}
        columns={columns}
        pageSize={8}
        rowsPerPageOptions={[5]}
        components={{ Toolbar: QuickSearchToolbar }}
        componentsProps={{
          toolbar: {
            value: searchText,
            onChange: (event) => requestSearch(event.target.value),
            clearSearch: () => requestSearch(''),
          },
        }}
      />
    </div>
  )
}
