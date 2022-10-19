import React, { Component } from 'react'
import { useState, useEffect, useRef } from 'react'
import { IconButton } from '@fluentui/react/lib/Button'
import {
  getTheme,
  mergeStyleSets,
  FontWeights,
  ContextualMenu,
  Toggle,
  Stack,
  Modal,
  IStackProps,
  IStackTokens,
  createTheme,
  loadTheme,
} from '@fluentui/react'
import { TextField } from '@fluentui/react/lib/TextField'
import { Label } from '@fluentui/react/lib/Label'
import {
  PrimaryButton,
  IButtonStyles,
  DefaultButton,
} from '@fluentui/react/lib/Button'
import { SearchBox, ISearchBoxStyles } from '@fluentui/react/lib/SearchBox'
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn,
} from '@fluentui/react/lib/DetailsList'
import {
  PeoplePicker,
  PrincipalType,
} from '@pnp/spfx-controls-react/lib/PeoplePicker'
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog'

import Pagination from 'office-ui-fabric-react-pagination'
//declare var ExcelJS;
import '../../../ExternalRef/workbench.css'
import * as FileSaver from 'file-saver'

import styles from './DeviceList.module.scss'
import ExcelJS from '../../../../node_modules/exceljs/dist/exceljs.min.js'
// import { loadTheme } from "office-ui-fabric-react";
let currentpage = 1
var totalPage = 30
const blueTheme = createTheme({
  palette: {
    themePrimary: '#004fa2',
    themeLighterAlt: '#f1f6fb',
    themeLighter: '#cadcf0',
    themeLight: '#9fc0e3',
    themeTertiary: '#508ac8',
    themeSecondary: '#155fae',
    themeDarkAlt: '#004793',
    themeDark: '#003c7c',
    themeDarker: '#002c5b',
    neutralLighterAlt: '#faf9f8',
    neutralLighter: '#f3f2f1',
    neutralLight: '#edebe9',
    neutralQuaternaryAlt: '#e1dfdd',
    neutralQuaternary: '#d0d0d0',
    neutralTertiaryAlt: '#c8c6c4',
    neutralTertiary: '#a19f9d',
    neutralSecondary: '#605e5c',
    neutralPrimaryAlt: '#3b3a39',
    neutralPrimary: '#323130',
    neutralDark: '#201f1e',
    black: '#000000',
    white: '#ffffff',
  },
})
const redTheme = createTheme({
  palette: {
    themePrimary: '#d71e2b',
    themeLighterAlt: '#fdf5f5',
    themeLighter: '#f8d6d9',
    themeLight: '#f3b4b8',
    themeTertiary: '#e77078',
    themeSecondary: '#db3540',
    themeDarkAlt: '#c11b26',
    themeDark: '#a31720',
    themeDarker: '#781118',
    neutralLighterAlt: '#faf9f8',
    neutralLighter: '#f3f2f1',
    neutralLight: '#edebe9',
    neutralQuaternaryAlt: '#e1dfdd',
    neutralQuaternary: '#d0d0d0',
    neutralTertiaryAlt: '#c8c6c4',
    neutralTertiary: '#a19f9d',
    neutralSecondary: '#605e5c',
    neutralPrimaryAlt: '#3b3a39',
    neutralPrimary: '#323130',
    neutralDark: '#201f1e',
    black: '#000000',
    white: '#ffffff',
  },
})
let allDeviceItems = []
const App = (props) => {
  const messagesEndRef = useRef(null)
  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }

  const BackIcon = () => (
    <IconButton
      iconProps={{
        iconName: 'NavigateBack',
        style: {
          fontSize: 35,
          color: requestType == 'wf' ? '#d71e2b' : '#004fa2',
        },
      }}
      title="Back"
      href={
        props.context.pageContext.web.absoluteUrl +
        `/SitePages/InternalForm.aspx?RequestID=${requestID}&RequestType=${requestType}`
      }
    />
  )
  const cancelIcon = { iconName: 'Cancel' }

  const stackTokens: IStackTokens = { childrenGap: 10 }
  const searchBoxStyles: Partial<ISearchBoxStyles> = {
    root: { width: 300, marginRight: 10 },
  }
  const paramsString = window.location.href.split('?')[1].toLowerCase()
  let requestType = ''
  let requestID = 0
  const PrimBtnStyles: Partial<IButtonStyles> = {
    root: {
      // backgroundColor: "#004fa2",
    },
  }

  const searchParams = new URLSearchParams(paramsString)
  searchParams.has('requesttype')
    ? (requestType = searchParams.get('requesttype'))
    : ''
  searchParams.has('requestid')
    ? (requestID = Number(searchParams.get('requestid')))
    : ''
  console.log(requestType)

  requestType == 'wf' ? loadTheme(redTheme) : loadTheme(blueTheme)

  const iconButtonStyles: Partial<IButtonStyles> = {
    root: {
      marginLeft: 'auto',
      marginTop: '4px',
      marginRight: '2px',
    },
  }
  const theme = getTheme()
  const dialogContentProps = {
    type: DialogType.normal,
    title: 'Are you sure you want to delete?',
    closeButtonAriaLabel: 'Close',
    subText: '',
  }
  const contentStyles = mergeStyleSets({
    container: {
      display: 'flex',
      flexFlow: 'column nowrap',
      alignItems: 'stretch',
      minWidth: 700,
      height: 625,
    },
    header: [
      // eslint-disable-next-line deprecation/deprecation
      theme.fonts.xLargePlus,
      {
        flex: '1 1 auto',
        // borderTop: `4px solid ${theme.palette.themePrimary}`,
        //color:#004fa2,
        display: 'flex',
        alignItems: 'center',
        fontWeight: FontWeights.semibold,
        padding: '12px 12px 14px 24px',
      },
    ],
    body: {
      flex: '4 4 auto',
      padding: '0 24px 24px 24px',
      overflowY: 'hidden',
      selectors: {
        p: { margin: '14px 0' },
        'p:first-child': { marginTop: 0 },
        'p:last-child': { marginBottom: 0 },
      },
    },
  })
  const importcontentStyles = mergeStyleSets({
    container: {
      display: 'flex',
      flexFlow: 'column nowrap',
      alignItems: 'stretch',
      width: 200,
      height: 150,
    },
    header: [
      // eslint-disable-next-line deprecation/deprecation
      theme.fonts.xLargePlus,
      {
        flex: '1 1 auto',
        // borderTop: `4px solid ${theme.palette.themePrimary}`,
        //color:#004fa2,
        display: 'flex',
        alignItems: 'center',
        fontWeight: FontWeights.semibold,
        padding: '12px 12px 14px 24px',
      },
    ],
    body: {
      flex: '4 4 auto',
      padding: '0 24px 24px 24px',
      overflowY: 'hidden',
      selectors: {
        p: { margin: '14px 0' },
        'p:first-child': { marginTop: 0 },
        'p:last-child': { marginBottom: 0 },
      },
    },
  })
  const [masterdeviceItems, setmasterdeviceItems] = useState([])
  const [deviceItems, setdeviceItems] = useState([])
  const [reRender, setReRender] = useState(true)
  const [isModalOpen, setModal] = useState(false)
  const [isNewItem, setNewItem] = useState(false)
  const [isDeleteConfirm, setDeleteConfirm] = useState(true)
  const [isImportModalOpen, setImportModalOpen] = useState(false)
  // const [viewItem, setViewItem] = useState({Device2ndSubnet:"",DevicePrimaryIP:"",DeviceSecondaryIP:"",DeviceType:"",Gateway:"",Id:0,Subnet:"",SupervisorIP:"",DeviceEdgeModel:"",DeviceHostID:"",DeviceSerialNumber:"",MACAddressEth0Port1:"",MACAddressEth1Port2:"",PanelType:"",SuperVisorNameId:0,SupervisorNameEmail:""});

  const [viewItem, setViewItem] = useState({
    Device2ndSubnet: '',
    DevicePrimaryIP: '',
    DeviceSecondaryIP: '',
    DeviceType: '',
    Gateway: '',
    Id: 0,
    Subnet: '',
    SupervisorIP: '',
    DeviceEdgeModel: '',
    DeviceHostID: '',
    DeviceSerialNumber: '',
    MACAddressEth0Port1: '',
    MACAddressEth1Port2: '',
    PanelType: '',
    SupervisorName: '',
    PrimaryMAC: '',
    SecondaryMAC: '',
    NiagaraVersion: '',
    Model: '',
    HostID: '',
    SerialNo: '',
    DateCode: '',
  })

  const _deviceColumns = [
    {
      key: 'devicetype',
      name: 'Device Type',
      fieldName: 'DeviceType',
      minWidth: 60,
      maxWidth: 120,
      isResizable: true,
      isRowHeader: true,
    },
    {
      key: 'deviceprimaryip',
      name: 'Device Primary IP',
      fieldName: 'DevicePrimaryIP',
      minWidth: 60,
      maxWidth: 120,
      isResizable: true,
      isRowHeader: true,
    },
    {
      key: 'subnet',
      name: 'Subnet',
      fieldName: 'Subnet',
      minWidth: 60,
      maxWidth: 100,
      isResizable: true,
      isRowHeader: true,
    },

    {
      key: 'gateway',
      name: 'Gateway',
      fieldName: 'Gateway',
      minWidth: 60,
      maxWidth: 100,
      isResizable: true,
      isRowHeader: true,
    },
    {
      key: 'devicesecondaryip',
      name: 'Device Secondary IP',
      fieldName: 'DeviceSecondaryIP',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      isRowHeader: true,
    },
    {
      key: 'device2ndsubnet',
      name: 'Device 2nd Subnet',
      fieldName: 'Device2ndSubnet',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      isRowHeader: true,
    },
    {
      key: 'supervisorip',
      name: 'Supervisor IP',
      fieldName: 'SupervisorIP',
      minWidth: 60,
      maxWidth: 100,
      isResizable: true,
      isRowHeader: true,
    },
    {
      key: 'supervisorname',
      name: 'Supervisor Name',
      fieldName: 'SupervisorName',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      isRowHeader: true,
    },
    {
      key: 'Actions',
      name: 'Actions',
      fieldName: 'Actions',
      minWidth: 190,
      maxWidth: 240,
      isResizable: true,
      isRowHeader: true,

      onRender: (item) => (
        <>
          <IconButton
            id={item.Id}
            iconProps={{
              iconName: 'EntryView',
              style: {
                fontSize: 20,
                color: requestType == 'wf' ? '#d71e2b' : '#004fa2',
              },
            }}
            title="View Form"
            onClick={(e) => modalFunction(e)}
          />
          <IconButton
            id={item.Id}
            iconProps={{
              iconName: 'Delete',
              style: {
                fontSize: 20,
                color: requestType == 'wf' ? '#d71e2b' : '#004fa2',
              },
            }}
            title="Delte"
            onClick={(e) => {
              setViewItem(
                deviceItems.filter(
                  (data) => data.Id == Number(e.currentTarget['id']),
                )[0],
              )
              setDeleteConfirm(false)
            }}
          />
        </>
      ),
    },
  ]

  useEffect(() => {
    if (reRender) {
      // props.spcontext.web.lists.getByTitle("DevicesList").items.select("*,SupervisorName/Title,SupervisorName/EMail,SupervisorName/ID").expand("SupervisorName").filter(`ReferenceID eq '${requestID}' and RecordType eq '${requestType}'`).orderBy("Created", false).get().then(async (deviceData: any)=>{
      props.spcontext.web.lists
        .getByTitle('DevicesList')
        .items.select(
          '*,SupervisorName/Title,SupervisorName/EMail,SupervisorName/ID',
        )
        .expand('SupervisorName')
        .filter(
          `ReferenceID eq '${requestID}' and RecordType eq '${requestType}'`,
        )
        .orderBy('Created', false)
        .get()
        .then(async (deviceData: any) => {
          allDeviceItems = []
          deviceData.forEach(async (dData) => {
            allDeviceItems.push({
              DeviceType: dData.DeviceType,
              DevicePrimaryIP: dData.DevicePrimaryIP,
              Subnet: dData.Subnet,
              Gateway: dData.Gateway,
              DeviceSecondaryIP: dData.DeviceSecondaryIP,
              Device2ndSubnet: dData.Device2ndSubnet,
              SupervisorIP: dData.SupervisorIP,
              // SupervisorNameEmail:dData.SupervisorName?dData.SupervisorName.EMail:"",
              // SupervisorNameId:dData.SupervisorName?dData.SupervisorName.Id:"",
              SupervisorName: dData.SupervisorNameText
                ? dData.SupervisorNameText
                : '',
              Id: dData.Id,
              DeviceEdgeModel: dData.DeviceEdgeModel,
              DeviceHostID: dData.DeviceHostID,
              DeviceSerialNumber: dData.DeviceSerialNumber,
              MACAddressEth0Port1: dData.MacAddressEth0Port1,
              MACAddressEth1Port2: dData.MacAddressEth1Port2,
              PanelType: dData.PanelType,
              PrimaryMAC: dData.PrimaryMAC,
              SecondaryMAC: dData.SecondaryMAC,
              NiagaraVersion: dData.NiagaraVersion,
              Model: dData.Model,
              HostID: dData.HostID,
              SerialNo: dData.SerialNo,
              DateCode: dData.DateCode,
            })
          })
          setdeviceItems(allDeviceItems)
          setmasterdeviceItems(allDeviceItems)
          paginate(1)
        })
        .then(() => {
          setReRender(false)
        })
    }

    scrollToBottom()
  }, [reRender])

  async function exportExcel() {
    const workbook = new ExcelJS.Workbook()
    const worksheet = workbook.addWorksheet('Device List')

    let dobCol = worksheet.getRow(1)

    worksheet.columns = [
      { header: 'ItemId', key: 'Id', width: 25 },
      { header: 'Device Type', key: 'DeviceType', width: 25 },
      { header: 'Device Primary IP', key: 'DevicePrimaryIP', width: 25 },
      { header: 'Subnet', key: 'Subnet', width: 25 },
      { header: 'Gateway', key: 'Gateway', width: 25 },
      { header: 'Device Secondary IP', key: 'DeviceSecondaryIP', width: 25 },
      { header: 'Device 2nd Subnet', key: 'Device2ndSubnet', width: 25 },
      { header: 'Supervisor IP', key: 'SupervisorIP', width: 25 },
      { header: 'Supervisor Name', key: 'SupervisorNameEmail', width: 25 },
      { header: 'Device/EdgeModel', key: 'DeviceEdgeModel', width: 25 },
      { header: 'Device Host ID', key: 'DeviceHostID', width: 25 },
      { header: 'Device Serial Number', key: 'DeviceSerialNumber', width: 25 },
      {
        header: 'MacAddress Eth0/Port1',
        key: 'MACAddressEth0Port1',
        width: 25,
      },
      {
        header: 'MacAddress Eth1/Port2',
        key: 'MACAddressEth1Port2',
        width: 25,
      },
      { header: 'PanelType', key: 'PanelType', width: 25 },
      { header: 'Primary MAC', key: 'PrimaryMAC', width: 25 },
      { header: 'Secondary MAC', key: 'SecondaryMAC', width: 25 },
      { header: 'Niagara Version', key: 'NiagaraVersion', width: 25 },
      { header: 'Model', key: 'Model', width: 25 },
      { header: 'Host ID', key: 'HostID', width: 25 },
      { header: 'Serial No.', key: 'SerialNo', width: 25 },
      { header: 'DateCode', key: 'DateCode', width: 25 },
    ]

    deviceItems.forEach(function (item, index) {
      worksheet.addRow({
        DeviceType: item.DeviceType,
        DevicePrimaryIP: item.DevicePrimaryIP,
        Subnet: item.Subnet,
        Gateway: item.Gateway,
        DeviceSecondaryIP: item.DeviceSecondaryIP,
        Device2ndSubnet: item.Device2ndSubnet,
        SupervisorIP: item.SupervisorIP,
        // SupervisorNameEmail:item.SupervisorNameEmail,
        SupervisorNameEmail: item.SupervisorName,
        Id: item.Id,
        DeviceEdgeModel: item.DeviceEdgeModel,
        DeviceHostID: item.DeviceHostID,
        DeviceSerialNumber: item.DeviceSerialNumber,
        MACAddressEth0Port1: item.MACAddressEth0Port1,
        MACAddressEth1Port2: item.MACAddressEth1Port2,
        PanelType: item.PanelType,
        PrimaryMAC: item.PrimaryMAC,
        SecondaryMAC: item.SecondaryMAC,
        NiagaraVersion: item.NiagaraVersion,
        Model: item.Model,
        HostID: item.HostID,
        SerialNo: item.SerialNo,
        DateCode: item.DateCode,
      })
    })
    ;[
      'A1',
      'B1',
      'C1',
      'D1',
      'E1',
      'F1',
      'G1',
      'H1',
      'I1',
      'J1',
      'K1',
      'L1',
      'M1',
      'N1',
      'O1',
      'P1',
      'Q1',
      'R1',
      'S1',
      'T1',
      'U1',
      'V1',
    ].map((key) => {
      worksheet.getCell(key).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'F2F2F2' },
      }
    })
    worksheet.eachRow({ includeEmpty: true }, function (cell, index) {
      cell._cells.map((key, index) => {
        worksheet.getCell(key._address).border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        }
      })
    })
    workbook.xlsx
      .writeBuffer()
      .then((buffer) =>
        FileSaver.saveAs(new Blob([buffer]), 'Device List' + '.xlsx'),
      )
      .catch((err) => console.log('Error writing excel export', err))
  }

  async function importExcel() {
    let rowArr = []
    let list = props.spcontext.web.lists.getByTitle('DevicesList')
    let createBatch = props.spcontext.web.createBatch()
    let updateBatch = props.spcontext.web.createBatch()
    const wb = new ExcelJS.Workbook()
    const reader = new FileReader()
    let filecontent = document.getElementById('uploadfile') as any
    await reader.readAsArrayBuffer(filecontent.files[0])
    reader.onload = () => {
      var buffer: any
      buffer = reader.result
      wb.xlsx
        .load(buffer)
        .then((workbook) => {
          workbook.eachSheet((sheet, id) => {
            sheet.eachRow(async (row, rowIndex) => {
              rowArr.push(row.values)
              if (rowIndex > 1) {
                let ItemId = row.values[rowArr[0].indexOf('ItemId')]
                if (!ItemId) {
                  list.items
                    .inBatch(createBatch)
                    .add({
                      DeviceType: row.values[
                        rowArr[0].indexOf('Device Type')
                      ].toString(),
                      DevicePrimaryIP: row.values[
                        rowArr[0].indexOf('Device Primary IP')
                      ].toString(),
                      Subnet: row.values[
                        rowArr[0].indexOf('Subnet')
                      ].toString(),
                      Gateway: row.values[
                        rowArr[0].indexOf('Gateway')
                      ].toString(),
                      DeviceSecondaryIP: row.values[
                        rowArr[0].indexOf('Device Secondary IP')
                      ].toString(),
                      Device2ndSubnet: row.values[
                        rowArr[0].indexOf('Device 2nd Subnet')
                      ].toString(),
                      SupervisorIP: row.values[
                        rowArr[0].indexOf('Supervisor IP')
                      ].toString(),
                      // SupervisorNameId:viewItem.SuperVisorNameId,
                      SupervisorNameText:
                        row.values[rowArr[0].indexOf('Supervisor Name')],
                      DeviceEdgeModel: row.values[
                        rowArr[0].indexOf('Device/EdgeModel')
                      ].toString(),
                      DeviceHostID: row.values[
                        rowArr[0].indexOf('Device Host ID')
                      ].toString(),
                      DeviceSerialNumber: row.values[
                        rowArr[0].indexOf('Device Serial Number')
                      ].toString(),
                      MacAddressEth0Port1: row.values[
                        rowArr[0].indexOf('MacAddress Eth0/Port1')
                      ].toString(),
                      MacAddressEth1Port2: row.values[
                        rowArr[0].indexOf('MacAddress Eth1/Port2')
                      ].toString(),
                      PanelType: row.values[
                        rowArr[0].indexOf('PanelType')
                      ].toString(),
                      PrimaryMAC: row.values[
                        rowArr[0].indexOf('Primary MAC')
                      ].toString(),
                      SecondaryMAC: row.values[
                        rowArr[0].indexOf('Secondary MAC')
                      ].toString(),
                      NiagaraVersion: row.values[
                        rowArr[0].indexOf('Niagara Version')
                      ].toString(),
                      Model: row.values[rowArr[0].indexOf('Model')].toString(),
                      HostID: row.values[rowArr[0].indexOf('HostID')],
                      SerialNo: row.values[
                        rowArr[0].indexOf('Serial No.')
                      ].toString(),
                      DateCode: row.values[
                        rowArr[0].indexOf('DateCode')
                      ].toString(),
                      ReferenceID: requestID.toString(),
                      RecordType: requestType == 'nwf' ? 'NWF' : 'WF',
                    })
                    .then((b) => {
                      console.log(b)
                    })
                } else {
                  list.items
                    .getById(ItemId)
                    .inBatch(updateBatch)
                    .update({
                      DeviceType: row.values[
                        rowArr[0].indexOf('Device Type')
                      ].toString(),
                      DevicePrimaryIP: row.values[
                        rowArr[0].indexOf('Device Primary IP')
                      ].toString(),
                      Subnet: row.values[
                        rowArr[0].indexOf('Subnet')
                      ].toString(),
                      Gateway: row.values[
                        rowArr[0].indexOf('Gateway')
                      ].toString(),
                      DeviceSecondaryIP: row.values[
                        rowArr[0].indexOf('Device Secondary IP')
                      ].toString(),
                      Device2ndSubnet: row.values[
                        rowArr[0].indexOf('Device 2nd Subnet')
                      ].toString(),
                      SupervisorIP: row.values[
                        rowArr[0].indexOf('Supervisor IP')
                      ].toString(),
                      // SupervisorNameId:viewItem.SuperVisorNameId,
                      SupervisorNameText:
                        row.values[rowArr[0].indexOf('SuperVisor Name')],
                      DeviceEdgeModel: row.values[
                        rowArr[0].indexOf('Device/EdgeModel')
                      ].toString(),
                      DeviceHostID: row.values[
                        rowArr[0].indexOf('Device Host ID')
                      ].toString(),
                      DeviceSerialNumber: row.values[
                        rowArr[0].indexOf('Device Serial Number')
                      ].toString(),
                      MacAddressEth0Port1: row.values[
                        rowArr[0].indexOf('MacAddress Eth0/Port1')
                      ].toString(),
                      MacAddressEth1Port2: row.values[
                        rowArr[0].indexOf('MacAddress Eth1/Port2')
                      ].toString(),
                      PanelType: row.values[
                        rowArr[0].indexOf('PanelType')
                      ].toString(),
                      PrimaryMAC: row.values[
                        rowArr[0].indexOf('Primary MAC')
                      ].toString(),
                      SecondaryMAC: row.values[
                        rowArr[0].indexOf('Secondary MAC')
                      ].toString(),
                      NiagaraVersion: row.values[
                        rowArr[0].indexOf('Niagara Version')
                      ].toString(),
                      Model: row.values[rowArr[0].indexOf('Model')].toString(),
                      HostID: row.values[rowArr[0].indexOf('HostID')],
                      SerialNo: row.values[
                        rowArr[0].indexOf('Serial No.')
                      ].toString(),
                      DateCode: row.values[
                        rowArr[0].indexOf('DateCode')
                      ].toString(),
                      ReferenceID: requestID.toString(),
                      RecordType: requestType == 'nwf' ? 'NWF' : 'WF',
                    })
                    .then((b) => {
                      console.log(b)
                    })
                }
              }
            })
          })
        })
        .then(async () => {
          if (createBatch._deps.length > 0) await createBatch.execute()
          if (updateBatch._deps.length > 0) await updateBatch.execute()

          setImportModalOpen(false)
          setModal(false)
          setReRender(true)
        })
    }
  }
  async function modalFunction(item) {
    setViewItem(
      deviceItems.filter((data) => data.Id == Number(item.currentTarget.id))[0],
    )
    setNewItem(false)
    setModal(true)
  }

  async function handleChange(newValue, params) {
    if (params == 'DevicePrimaryIP')
      setViewItem({ ...viewItem, DevicePrimaryIP: newValue })
    else if (params == 'DeviceType')
      setViewItem({ ...viewItem, DeviceType: newValue })
    else if (params == 'Subnet') setViewItem({ ...viewItem, Subnet: newValue })
    else if (params == 'DeviceEdgeModel')
      setViewItem({ ...viewItem, DeviceEdgeModel: newValue })
    else if (params == 'Gateway')
      setViewItem({ ...viewItem, Gateway: newValue })
    else if (params == 'DeviceHostID')
      setViewItem({ ...viewItem, DeviceHostID: newValue })
    else if (params == 'DeviceSecondaryIP')
      setViewItem({ ...viewItem, DeviceSecondaryIP: newValue })
    else if (params == 'DeviceSerialNumber')
      setViewItem({ ...viewItem, DeviceSerialNumber: newValue })
    else if (params == 'Device2ndSubnet')
      setViewItem({ ...viewItem, Device2ndSubnet: newValue })
    else if (params == 'MACAddressEth0Port1')
      setViewItem({ ...viewItem, MACAddressEth0Port1: newValue })
    else if (params == 'MACAddressEth1Port2')
      setViewItem({ ...viewItem, MACAddressEth1Port2: newValue })
    else if (params == 'SupervisorIP')
      setViewItem({ ...viewItem, SupervisorIP: newValue })
    else if (params == 'SupervisorName')
      setViewItem({ ...viewItem, SupervisorName: newValue })
    // else if(params=="SupervisorName"&&newValue.length>0)
    // setViewItem({...viewItem, SupervisorNameEmail:newValue[0].secondaryText,SuperVisorNameId:newValue[0].id});
    else if (params == 'PanelType')
      setViewItem({ ...viewItem, PanelType: newValue })
    else if (params == 'PrimaryMAC')
      setViewItem({ ...viewItem, PrimaryMAC: newValue })
    else if (params == 'SecondaryMAC')
      setViewItem({ ...viewItem, SecondaryMAC: newValue })
    else if (params == 'NiagaraVersion')
      setViewItem({ ...viewItem, NiagaraVersion: newValue })
    else if (params == 'Model') setViewItem({ ...viewItem, Model: newValue })
    else if (params == 'HostID') setViewItem({ ...viewItem, HostID: newValue })
    else if (params == 'SerialNo')
      setViewItem({ ...viewItem, SerialNo: newValue })
    else if (params == 'DateCode')
      setViewItem({ ...viewItem, DateCode: newValue })
  }

  async function DeleteItem() {
    await props.spcontext.web.lists
      .getByTitle('DevicesList')
      .items.getById(viewItem.Id)
      .delete()
      .then(() => {
        setReRender(true)
        setViewItem({
          Device2ndSubnet: '',
          DevicePrimaryIP: '',
          DeviceSecondaryIP: '',
          DeviceType: '',
          Gateway: '',
          Id: 0,
          Subnet: '',
          SupervisorIP: '',
          DeviceEdgeModel: '',
          DeviceHostID: '',
          DeviceSerialNumber: '',
          MACAddressEth0Port1: '',
          MACAddressEth1Port2: '',
          PanelType: '',
          SupervisorName: '',
          PrimaryMAC: '',
          SecondaryMAC: '',
          NiagaraVersion: '',
          Model: '',
          HostID: '',
          SerialNo: '',
          DateCode: '',
        })
        setDeleteConfirm(true)
      })
  }

  async function updateDeviceItem(ItemID) {
    if (!isNewItem) {
      await props.spcontext.web.lists
        .getByTitle('DevicesList')
        .items.getById(viewItem.Id)
        .update({
          DeviceType: viewItem.DeviceType,
          DevicePrimaryIP: viewItem.DevicePrimaryIP,
          Subnet: viewItem.Subnet,
          Gateway: viewItem.Gateway,
          DeviceSecondaryIP: viewItem.DeviceSecondaryIP,
          Device2ndSubnet: viewItem.Device2ndSubnet,
          SupervisorIP: viewItem.SupervisorIP,
          // SupervisorNameId:viewItem.SuperVisorNameId,
          SupervisorNameText: viewItem.SupervisorName,
          DeviceEdgeModel: viewItem.DeviceEdgeModel,
          DeviceHostID: viewItem.DeviceHostID,
          DeviceSerialNumber: viewItem.DeviceSerialNumber,
          MacAddressEth0Port1: viewItem.MACAddressEth0Port1,
          MacAddressEth1Port2: viewItem.MACAddressEth1Port2,
          PanelType: viewItem.PanelType,
          PrimaryMAC: viewItem.PrimaryMAC,
          SecondaryMAC: viewItem.SecondaryMAC,
          NiagaraVersion: viewItem.NiagaraVersion,
          Model: viewItem.Model,
          HostID: viewItem.HostID,
          SerialNo: viewItem.SerialNo,
          DateCode: viewItem.DateCode,
        })
        .then(() => {
          setReRender(true)
          setModal(false)
          // setViewItem({Device2ndSubnet:"",DevicePrimaryIP:"",DeviceSecondaryIP:"",DeviceType:"",Gateway:"",Id:0,Subnet:"",SupervisorIP:"",DeviceEdgeModel:"",DeviceHostID:"",DeviceSerialNumber:"",MACAddressEth0Port1:"",MACAddressEth1Port2:"",PanelType:"",SuperVisorNameId:0,SupervisorNameEmail:""});
          setViewItem({
            Device2ndSubnet: '',
            DevicePrimaryIP: '',
            DeviceSecondaryIP: '',
            DeviceType: '',
            Gateway: '',
            Id: 0,
            Subnet: '',
            SupervisorIP: '',
            DeviceEdgeModel: '',
            DeviceHostID: '',
            DeviceSerialNumber: '',
            MACAddressEth0Port1: '',
            MACAddressEth1Port2: '',
            PanelType: '',
            SupervisorName: '',
            PrimaryMAC: '',
            SecondaryMAC: '',
            NiagaraVersion: '',
            Model: '',
            HostID: '',
            SerialNo: '',
            DateCode: '',
          })
        })
    } else {
      await props.spcontext.web.lists
        .getByTitle('DevicesList')
        .items.add({
          ReferenceID: requestID.toString(),
          RecordType: requestType.toUpperCase(),
          DeviceType: viewItem.DeviceType,
          DevicePrimaryIP: viewItem.DevicePrimaryIP,
          Subnet: viewItem.Subnet,
          Gateway: viewItem.Gateway,
          DeviceSecondaryIP: viewItem.DeviceSecondaryIP,
          Device2ndSubnet: viewItem.Device2ndSubnet,
          SupervisorIP: viewItem.SupervisorIP,
          // SupervisorNameId:viewItem.SuperVisorNameId,
          SupervisorNameText: viewItem.SupervisorName,
          DeviceEdgeModel: viewItem.DeviceEdgeModel,
          DeviceHostID: viewItem.DeviceHostID,
          DeviceSerialNumber: viewItem.DeviceSerialNumber,
          MacAddressEth0Port1: viewItem.MACAddressEth0Port1,
          MacAddressEth1Port2: viewItem.MACAddressEth1Port2,
          PanelType: viewItem.PanelType,
          PrimaryMAC: viewItem.PrimaryMAC,
          SecondaryMAC: viewItem.SecondaryMAC,
          NiagaraVersion: viewItem.NiagaraVersion,
          Model: viewItem.Model,
          HostID: viewItem.HostID,
          SerialNo: viewItem.SerialNo,
          DateCode: viewItem.DateCode,
        })
        .then(() => {
          setReRender(true)
          setModal(false)
          setNewItem(false)
          // setViewItem({Device2ndSubnet:"",DevicePrimaryIP:"",DeviceSecondaryIP:"",DeviceType:"",Gateway:"",Id:0,Subnet:"",SupervisorIP:"",DeviceEdgeModel:"",DeviceHostID:"",DeviceSerialNumber:"",MACAddressEth0Port1:"",MACAddressEth1Port2:"",PanelType:"",SuperVisorNameId:0,SupervisorNameEmail:""});
          setViewItem({
            Device2ndSubnet: '',
            DevicePrimaryIP: '',
            DeviceSecondaryIP: '',
            DeviceType: '',
            Gateway: '',
            Id: 0,
            Subnet: '',
            SupervisorIP: '',
            DeviceEdgeModel: '',
            DeviceHostID: '',
            DeviceSerialNumber: '',
            MACAddressEth0Port1: '',
            MACAddressEth1Port2: '',
            PanelType: '',
            SupervisorName: '',
            PrimaryMAC: '',
            SecondaryMAC: '',
            NiagaraVersion: '',
            Model: '',
            HostID: '',
            SerialNo: '',
            DateCode: '',
          })
        })
    }
  }
  const paginate = (pagenumber) => {
    var lastIndex = pagenumber * totalPage
    var firstIndex = lastIndex - totalPage
    var paginatedItems = allDeviceItems.slice(firstIndex, lastIndex)
    currentpage = pagenumber
    setdeviceItems([...paginatedItems])
  }
  return (
    <>
      <div
        style={{
          display: 'flex',
          color: requestType == 'wf' ? '#d71e2b' : '#004fa2',
        }}
      >
        <BackIcon />
        <Label
          style={{
            fontSize: 25,
            margin: 'auto',
            color: requestType == 'wf' ? '#d71e2b' : '#004fa2',
          }}
        >
          Devices
        </Label>
      </div>
      <div>
        {/* {messages.map(message => <Message key={message.id} {...message} />)} */}
        <div ref={messagesEndRef} />
      </div>
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          marginTop: '30px',
        }}
      >
        <div style={{ display: 'flex', marginBottom: '30px' }}>
          <Stack horizontal tokens={stackTokens} style={{ display: 'flex' }}>
            <PrimaryButton
              text="Import Excel"
              onClick={() => {
                setImportModalOpen(true)
                setModal(false)
              }}
              style={{ marginLeft: '20px' }}
            />
            <PrimaryButton
              text="Export Excel"
              href={
                deviceItems.length == 0
                  ? 'https://chandrudemo.sharepoint.com/sites/LynxSpring/Shared%20Documents/Sample Device List.xlsx'
                  : ''
              }
              onClick={exportExcel}
            />
            <PrimaryButton
              text="Add New"
              onClick={() => {
                setModal(true)
                setNewItem(true)
              }}
            />
          </Stack>
        </div>
        <div style={{ display: 'flex' }}>
          <SearchBox
            className="SearchBox"
            placeholder="Search ..."
            styles={searchBoxStyles}
            onChange={(_, newValue) => {
              if (newValue) {
                setdeviceItems(
                  masterdeviceItems.filter((item) =>
                    item.DeviceType.toLowerCase().includes(
                      newValue.toLowerCase(),
                    ),
                  ),
                )
              } else {
                setdeviceItems(masterdeviceItems)
              }
            }}
            onSearch={(newValue) => {
              {
                if (newValue) {
                  setdeviceItems(
                    masterdeviceItems.filter((item) =>
                      item.DeviceType.toLowerCase().includes(
                        newValue.toLowerCase(),
                      ),
                    ),
                  )
                } else {
                  setdeviceItems(masterdeviceItems)
                }
              }
            }}
          />
        </div>
      </div>
      {deviceItems.length > 0 ? (
        <div className={styles.DetailsListSection}>
          <Pagination
            style={{ margin: 'auto' }}
            currentPage={currentpage}
            totalPages={
              allDeviceItems.length > 0
                ? Math.ceil(allDeviceItems.length / totalPage)
                : 1
            }
            onChange={(page) => {
              paginate(page)
            }}
          />
          <DetailsList
            items={deviceItems}
            columns={_deviceColumns}
            setKey="none"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selectionMode={SelectionMode.none}
          />
        </div>
      ) : (
        <div className={styles.noDataFound}>No Data Found</div>
      )}

      <Dialog
        hidden={isDeleteConfirm}
        onDismiss={() => {
          setDeleteConfirm(true)
        }}
        dialogContentProps={dialogContentProps}
      >
        <DialogFooter>
          <PrimaryButton onClick={() => DeleteItem()} text="Ok" />
          <DefaultButton
            text="Cancel"
            onClick={() => {
              setDeleteConfirm(true)
            }}
          />
        </DialogFooter>
      </Dialog>

      <Modal
        isOpen={isModalOpen}
        onDismiss={() => setModal(false)}
        isBlocking={false}
        containerClassName={contentStyles.container}
      >
        <div className={contentStyles.header}>
          <IconButton
            styles={iconButtonStyles}
            iconProps={cancelIcon}
            ariaLabel="Close popup modal"
            onClick={() => {
              setModal(false)
              setViewItem({
                Device2ndSubnet: '',
                DevicePrimaryIP: '',
                DeviceSecondaryIP: '',
                DeviceType: '',
                Gateway: '',
                Id: 0,
                Subnet: '',
                SupervisorIP: '',
                DeviceEdgeModel: '',
                DeviceHostID: '',
                DeviceSerialNumber: '',
                MACAddressEth0Port1: '',
                MACAddressEth1Port2: '',
                PanelType: '',
                SupervisorName: '',
                PrimaryMAC: '',
                SecondaryMAC: '',
                NiagaraVersion: '',
                Model: '',
                HostID: '',
                SerialNo: '',
                DateCode: '',
              })
            }}
          />
        </div>
        <div className={contentStyles.body}>
          <div className={styles.modalContent}>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Device Primary IP"
                  value={viewItem.DevicePrimaryIP}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'DevicePrimaryIP')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Device Type"
                  value={viewItem.DeviceType}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'DeviceType')
                  }}
                  className={styles.inputSectiontxtinput}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Subnet"
                  value={viewItem.Subnet}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'Subnet')
                  }}
                />
              </div>
            </div>
          </div>
          {/* <div style={{display: 'flex',padding:7}}>
            <div><TextField label="Device Primary IP" value={viewItem.DevicePrimaryIP} onChange={(e,newValue)=>{handleChange(newValue,"DevicePrimaryIP")}}/></div>
            <div><TextField label="Device Type" value={viewItem.DeviceType} onChange={(e,newValue)=>{handleChange(newValue,"DeviceType")}}/></div>
          </div> */}

          {/* <div style={{display: 'flex'}}>
            <div><TextField label="Subnet" value={viewItem.Subnet} onChange={(e,newValue)=>{handleChange(newValue,"Subnet")}}/></div>
            <div><TextField label="Device/Edge Model" value={viewItem.DeviceEdgeModel} onChange={(e,newValue)=>{handleChange(newValue,"DeviceEdgeModel")}}/></div>
          </div> */}
          <div className={styles.modalContent}>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Device/Edge Model"
                  value={viewItem.DeviceEdgeModel}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'DeviceEdgeModel')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Gateway"
                  value={viewItem.Gateway}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'Gateway')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Device Host ID"
                  value={viewItem.DeviceHostID}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'DeviceHostID')
                  }}
                />
              </div>
            </div>
          </div>
          {/* <div style={{display: 'flex'}}>
            <div><TextField label="Gateway" value={viewItem.Gateway} onChange={(e,newValue)=>{handleChange(newValue,"Gateway")}}/></div>
            <div><TextField label="Device Host ID" value={viewItem.DeviceHostID} onChange={(e,newValue)=>{handleChange(newValue,"DeviceHostID")}}/></div>
          </div> */}
          <div className={styles.modalContent}>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Device Secondary IP"
                  value={viewItem.DeviceSecondaryIP}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'DeviceSecondaryIP')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Device Serial #"
                  value={viewItem.DeviceSerialNumber}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'DeviceSerialNumber')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Device 2nd Subnet"
                  value={viewItem.Device2ndSubnet}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'Device2ndSubnet')
                  }}
                />
              </div>
            </div>
          </div>
          {/* <div style={{display: 'flex'}}>
            <div><TextField label="Device Secondary IP" value={viewItem.DeviceSecondaryIP} onChange={(e,newValue)=>{handleChange(newValue,"DeviceSecondaryIP")}}/></div>
            <div><TextField label="Device Serial #" value={viewItem.DeviceSerialNumber} onChange={(e,newValue)=>{handleChange(newValue,"DeviceSerialNumber")}}/></div>
          </div> */}
          {/* <div style={{display: 'flex'}}>
            <div><TextField label="Device 2nd Subnet" value={viewItem.Device2ndSubnet} onChange={(e,newValue)=>{handleChange(newValue,"Device2ndSubnet")}}/></div>
            <div><TextField label="MAC Address Eth0/Port1" value={viewItem.MACAddressEth0Port1} onChange={(e,newValue)=>{handleChange(newValue,"MACAddressEth0Port1")}}/></div>
          </div> */}
          <div className={styles.modalContent}>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="MAC Address Eth0/Port1"
                  value={viewItem.MACAddressEth0Port1}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'MACAddressEth0Port1')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Supervisor IP"
                  value={viewItem.SupervisorIP}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'SupervisorIP')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="MAC Address Eth1/Port2"
                  value={viewItem.MACAddressEth1Port2}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'MACAddressEth1Port2')
                  }}
                />
              </div>
            </div>
          </div>
          {/* <div style={{display: 'flex'}}>
            <div><TextField label="Supervisor IP" value={viewItem.SupervisorIP} onChange={(e,newValue)=>{handleChange(newValue,"SupervisorIP")}}/></div>
            <div><TextField label="MAC Address Eth1/Port2" value={viewItem.MACAddressEth1Port2} onChange={(e,newValue)=>{handleChange(newValue,"MACAddressEth1Port2")}}/></div>
          </div> */}
          {/* <div style={{display: 'flex'}}> */}
          <div className={styles.modalContent}>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Supervisor Name"
                  value={viewItem.SupervisorName}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'SupervisorName')
                  }}
                />
                {/* <PeoplePicker
              context={props.context}
              titleText="Supervisor Name"
              personSelectionLimit={1}
              groupName={""}
              showtooltip={true}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              resolveDelay={1000}
              defaultSelectedUsers={[viewItem.SupervisorNameEmail]}
              onChange={(e)=>{handleChange(e,"SupervisorName")}}
              ensureUser={true}/> */}
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Panel Type"
                  value={viewItem.PanelType}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'PanelType')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Primary MAC"
                  value={viewItem.PrimaryMAC}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'PrimaryMAC')
                  }}
                />
              </div>
            </div>
          </div>

          <div className={styles.modalContent}>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Secondary MAC"
                  value={viewItem.SecondaryMAC}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'SecondaryMAC')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Niagara Version"
                  value={viewItem.NiagaraVersion}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'NiagaraVersion')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Model"
                  value={viewItem.Model}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'Model')
                  }}
                />
              </div>
            </div>
          </div>

          <div className={styles.modalContent}>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="HostID"
                  value={viewItem.HostID}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'HostID')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="Serial No."
                  value={viewItem.SerialNo}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'SerialNo')
                  }}
                />
              </div>
            </div>
            <div className={styles.inputSection}>
              <div>
                <TextField
                  label="DateCode"
                  value={viewItem.DateCode}
                  onChange={(e, newValue) => {
                    handleChange(newValue, 'DateCode')
                  }}
                />
              </div>
            </div>
          </div>

          <div className={styles.viewformbtn}>
            <PrimaryButton
              styles={PrimBtnStyles}
              text="Submit"
              onClick={() => updateDeviceItem(viewItem.Id)}
              className={styles.submitbtn}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => {
                setModal(false)
                setViewItem({
                  Device2ndSubnet: '',
                  DevicePrimaryIP: '',
                  DeviceSecondaryIP: '',
                  DeviceType: '',
                  Gateway: '',
                  Id: 0,
                  Subnet: '',
                  SupervisorIP: '',
                  DeviceEdgeModel: '',
                  DeviceHostID: '',
                  DeviceSerialNumber: '',
                  MACAddressEth0Port1: '',
                  MACAddressEth1Port2: '',
                  PanelType: '',
                  SupervisorName: '',
                  PrimaryMAC: '',
                  SecondaryMAC: '',
                  NiagaraVersion: '',
                  Model: '',
                  HostID: '',
                  SerialNo: '',
                  DateCode: '',
                })
              }}
            />
          </div>
        </div>
      </Modal>

      <Modal
        isOpen={isImportModalOpen}
        onDismiss={() => {
          setModal(false)
          setImportModalOpen(false)
          setViewItem({
            Device2ndSubnet: '',
            DevicePrimaryIP: '',
            DeviceSecondaryIP: '',
            DeviceType: '',
            Gateway: '',
            Id: 0,
            Subnet: '',
            SupervisorIP: '',
            DeviceEdgeModel: '',
            DeviceHostID: '',
            DeviceSerialNumber: '',
            MACAddressEth0Port1: '',
            MACAddressEth1Port2: '',
            PanelType: '',
            SupervisorName: '',
            PrimaryMAC: '',
            SecondaryMAC: '',
            NiagaraVersion: '',
            Model: '',
            HostID: '',
            SerialNo: '',
            DateCode: '',
          })
        }}
        isBlocking={false}
        containerClassName={importcontentStyles.container}
      >
        <div className={importcontentStyles.header}>
          <IconButton
            styles={iconButtonStyles}
            iconProps={cancelIcon}
            ariaLabel="Close popup modal"
            onClick={() => {
              setModal(false)
              setImportModalOpen(false)
              setViewItem({
                Device2ndSubnet: '',
                DevicePrimaryIP: '',
                DeviceSecondaryIP: '',
                DeviceType: '',
                Gateway: '',
                Id: 0,
                Subnet: '',
                SupervisorIP: '',
                DeviceEdgeModel: '',
                DeviceHostID: '',
                DeviceSerialNumber: '',
                MACAddressEth0Port1: '',
                MACAddressEth1Port2: '',
                PanelType: '',
                SupervisorName: '',
                PrimaryMAC: '',
                SecondaryMAC: '',
                NiagaraVersion: '',
                Model: '',
                HostID: '',
                SerialNo: '',
                DateCode: '',
              })
            }}
          />
        </div>
        <div className={importcontentStyles.body}>
          <input
            placeholder="Please Select File"
            className={styles.customfileupload}
            type="file"
            id="uploadfile"
            accept=".xlsx"
          />
          <div
            style={{
              display: 'flex',
              justifyContent: 'center',
              marginTop: '20px',
            }}
          >
            <PrimaryButton
              text="Import"
              onClick={importExcel}
              className={styles.Importsubmitbtn}
            />
            <DefaultButton
              text="Cancel"
              onClick={() => {
                setModal(false)
                setViewItem({
                  Device2ndSubnet: '',
                  DevicePrimaryIP: '',
                  DeviceSecondaryIP: '',
                  DeviceType: '',
                  Gateway: '',
                  Id: 0,
                  Subnet: '',
                  SupervisorIP: '',
                  DeviceEdgeModel: '',
                  DeviceHostID: '',
                  DeviceSerialNumber: '',
                  MACAddressEth0Port1: '',
                  MACAddressEth1Port2: '',
                  PanelType: '',
                  SupervisorName: '',
                  PrimaryMAC: '',
                  SecondaryMAC: '',
                  NiagaraVersion: '',
                  Model: '',
                  HostID: '',
                  SerialNo: '',
                  DateCode: '',
                })
              }}
            />
          </div>
        </div>
      </Modal>
    </>
  )
}

export default App
