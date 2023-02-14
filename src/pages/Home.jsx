import { Button, Container, Link, Typography } from '@mui/material'
import React, { useEffect, useState } from 'react'
import { AlertDialog } from '../components/AlertDialog'
import { DataTable } from '../components/DataTable'
import { SimpleBackdrop } from '../components/SimpleBackdrop'




export const Home = () => {

  const [showBackdrop, setShowBackdrop] = useState(true)
  const [dialogData, setDialogData] = useState({title: "", message: "", open: false})
  const [originalRows, setOriginalRows] = useState()
  const [columns, setColumns] = useState(
    [
      { field: 'projNo', headerName: 'Project Number', minWidth: 130, flex: 0.5 },
      {
        field: 'projName', headerName: 'Project Name', minWidth: 200, flex: 1,
        renderCell: (params) => <Link href={"https://drive.google.com/drive/folders/" + params.row.id}  target="_blank">{ params.row.projName }</Link>
      },
    ]
  )

  const handleUnlock = (id) => {
    setShowBackdrop(true)
    //Unlock project -> success: remove row from table + shows success dialog || error: shows error message
    google.script.run
      .withSuccessHandler((url) => {
          setOriginalRows((prev) => prev.filter(row => row.id !== id))
          setDialogData({ title: "Unlocked successfully", message: <a href={url} target="_blank">Drive link</a>, open: true })
          setShowBackdrop(false)
        }
      )
      .withFailureHandler(onFailure)
      .moveLockedToDriveJ(id)
  }

  const handleCloseDialog = () => {
    setDialogData({title: "", message: "", open: false})
  };

  const onFailure = (error) => {
    console.log(error)
    setDialogData({ title: "Error!", message: error.message, open: true })
    setShowBackdrop(false)
  }

  useEffect(() => {
    //get all locked projects data from server
    google.script.run
      .withSuccessHandler((data) => {
        if (data) {
          let { hasPermission, projData } = data
          if (hasPermission) {
            setColumns((prev) => [...prev, {
              field: 'action',
              headerName: 'Action',
              description: 'This column is not sortable.',
              sortable: false,
              filterable: false,
              maxWidth: 160,
              flex: 0.5,
              disableExport: true, 
              renderCell: (params) => <Button variant='outlined' onClick={() => handleUnlock(params.row.id)}>Unlock folder</Button>
            }])
          }
          setOriginalRows(projData)
        }
        setShowBackdrop(false)
      })
      .withFailureHandler(onFailure)
      .getLockedDriveProjects()
  }, [])

  return (
    <Container>
      <Typography
        variant='h1'
        sx={{ my: 4, textAlign: "center", color: "primary.main" }}
      >
        Locked folders
      </Typography>
      <Typography variant='h2' sx={{ my: 2 }} >
        Locked projects list
      </Typography>

      {/* Render Datatable when data returned from server */}
      {originalRows && <DataTable columns={columns} originalRows={originalRows} />}
      {dialogData.open && <AlertDialog title={dialogData.title} message={dialogData.message} handleCloseDialog={handleCloseDialog} />}
      {showBackdrop && <SimpleBackdrop />}
    </Container>
  )
}
