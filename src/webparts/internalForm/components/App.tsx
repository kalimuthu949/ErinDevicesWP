import * as React from 'react'
import { useState, useEffect, useRef } from 'react'
import '../../../ExternalRef/workbench.css'
import styles from './InternalForm.module.scss'
import { Icon } from '@fluentui/react/lib/Icon'
import { IconButton } from '@fluentui/react/lib/Button'
import { TextField, MaskedTextField } from '@fluentui/react/lib/TextField'
import { Label } from '@fluentui/react/lib/Label'
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from '@fluentui/react/lib/ChoiceGroup'
import { DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button'
import { Checkbox, Stack, IIconProps, TextStyles } from '@fluentui/react'
import {
  PeoplePicker,
  PrincipalType,
} from '@pnp/spfx-controls-react/lib/PeoplePicker'
import { ThemeProvider, PartialTheme, createTheme } from '@fluentui/react'

import {
  DatePicker,
  DayOfWeek,
  Dropdown,
  IDropdownOption,
  mergeStyles,
  defaultDatePickerStrings,
} from '@fluentui/react'
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog'

const halfWidthInput = {
  root: { width: 300, margin: '0 1rem 0.5rem 0' },
}

interface Itaskdetails {
  TaskChecked: false
  TaskName: ''
  CompletedBy: ''
  Userid: ''
  Date: Date
}
;[]

interface IShippingdetails {
  ShippingDate: Date
  TrackingNumber: string
  CarrierNumber: string
}
;[]

var count = 0

const App = (props) => {
  // const messagesEndRef = useRef(null)
  // const scrollToBottom = () => {
  //   messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  // }
  let requestID = 0
  let requestType = ''
  let InternalItems = []
  const paramsString = window.location.href.split('?')[1].toLowerCase()
  const searchParams = new URLSearchParams(paramsString)
  searchParams.has('requestid')
    ? (requestID = Number(searchParams.get('requestid')))
    : ''
  searchParams.has('requesttype')
    ? (requestType = searchParams.get('requesttype'))
    : ''
  const dialogStyles = { main: { maxWidth: 450 } }
  const dialogContentProps = {
    type: DialogType.normal,
    title: '',
    closeButtonAriaLabel: 'Close',
    subText: 'Saved Successfully',
  }
  const fullWidthInput = {
    root: { width: '200px' },
  }
  // const halfWidthInput = {
  //   root: { width: "38%" },
  // };

  const onFormatDate = (date?: Date): string => {
    return date.getMonth() + 1 + '/' + date.getDate() + '/' + date.getFullYear()
  }

  const options: IChoiceGroupOption[] = [
    { key: 'A', text: 'Yes' },
    { key: 'B', text: 'No' },
  ]

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
  const choiceGroupStyles = {
    flexContainer: {
      display: 'flex',
      width: 300,
      margin: '0 1rem 0.5rem 0',
      label: {
        marginRight: '1rem',
      },
    },
  }
  const addIcon: IIconProps = { iconName: 'Add' }
  var tasks: Itaskdetails[] = []
  var shippingDetails: IShippingdetails[] = []
  const [firstDayOfWeek, setFirstDayOfWeek] = useState(DayOfWeek.Sunday)
  const [hideDialog, setHideDialog] = useState(true)
  const [rows, setrows] = useState(0)
  const [newtasks, setnewtasks] = useState(tasks)
  const [newShippingDetails, setnewShippingDetails] = useState(shippingDetails)
  const [internalFormItem, setinternalFormItem] = useState({
    BENumber: '',
    ProjectName: '',
    BuilderInitials: '',
    ProjectManager: '',
    ProjectDescription: '',
    Longtitude: '',
    Latitude: '',
    Address: '',
    State: '',
    Zipcode: '',
    DateofQuote: new Date(),
    DateQuoteSent: new Date(),
    ShippingDate: new Date(),
    TrackingNumber: '',
    POIssued: false,
    UtilityNetInformation: '',
    InternalConfig: false,
    InternalConfigAssignedToEmail: '',
    InternalConfigAssignedToId: 0,
    InternalConfigDate: new Date(),
    LightingLiveConfig: false,
    LightingLiveConfigAssignedToEmail: '',
    LightingLiveConfigAssignedToId: 0,
    LightingLiveConfigDate: new Date(),
    HVACConfig: false,
    HVACConfigAssignedToEmail: '',
    HVACConfigAssignedToId: 0,
    HVACConfigDate: new Date(),
    OtherNetworkIpsOnsite: false,
    OtherNetworkIpsAssignedToEmail: '',
    OtherNetworkIpsAssignedToId: 0,
    OtherNetworkIpsOnsiteDate: new Date(),
    StationName: false,
    StationNameAssignedToEmail: '',
    StationNameAssignedToId: 0,
    StationNameDate: new Date(),
    OnboardToSupervisor: false,
    OnboardAssignedToEmail: '',
    OnboardAssignedToId: 0,
    OnboardDate: new Date(),
    Id: 0,
    Notes: '',
  })
  const [reRender, setReRender] = useState(true)

  useEffect(() => {
    if (reRender) {
      if (requestType && requestID) {
        props.spcontext.web.lists
          .getByTitle('InternalFormList')
          .items.select(
            '*,InternalConfigAssignedTo/Title,InternalConfigAssignedTo/EMail,InternalConfigAssignedTo/ID,LightingLiveConfigAssignedTo/Title,LightingLiveConfigAssignedTo/EMail,LightingLiveConfigAssignedTo/ID,HVACConfigAssignedTo/Title,HVACConfigAssignedTo/EMail,HVACConfigAssignedTo/ID,OtherNetworkIpsAssignedTo/Title,OtherNetworkIpsAssignedTo/EMail,OtherNetworkIpsAssignedTo/ID,StationNameAssignedTo/Title,StationNameAssignedTo/EMail,StationNameAssignedTo/ID,OnboardAssignedTo/Title,OnboardAssignedTo/EMail,OnboardAssignedTo/ID',
          )
          .expand(
            'InternalConfigAssignedTo,LightingLiveConfigAssignedTo,HVACConfigAssignedTo,OtherNetworkIpsAssignedTo,StationNameAssignedTo,OnboardAssignedTo',
          )
          .filter(
            `ReferenceID eq '${requestID}' and RecordType eq '${requestType}'`,
          )
          .orderBy('Created', false)
          .get()
          .then(async (InternalData: any) => {
            if (InternalData.length > 0) {
              InternalData.forEach(async (dData) => {
                InternalItems.push({
                  BENumber: dData.BENumber,
                  ProjectName: dData.ProjectName,
                  BuilderInitials: dData.BuilderInitials,
                  ProjectManager: dData.ProjectManager,
                  ProjectDescription: dData.ProjectDescription,
                  Device2ndSubnet: dData.Device2ndSubnet,
                  Longtitude: dData.Longtitude,
                  Latitude: dData.Latitude,
                  Address: dData.StreetAddress,
                  State: dData.State,
                  Zipcode: dData.Zipcode,
                  DateofQuote: dData.DateofQuote,
                  DateQuoteSent: dData.DateQuoteSent,
                  TrackingNumber: dData.TrackingNumber,
                  ShippingDate: dData.ShippingDate,
                  POIssued: dData.POIssued,
                  UtilityNetInformation: dData.UtilityNetInformation,
                  SupervisorIP: dData.SupervisorIP,
                  InternalConfig: dData.InternalConfig,
                  InternalConfigAssignedToEmail: dData.InternalConfigAssignedTo
                    ? dData.InternalConfigAssignedTo.EMail
                    : '',
                  InternalConfigAssignedToId: dData.InternalConfigAssignedTo
                    ? dData.InternalConfigAssignedTo.ID
                    : '',
                  InternalConfigDate: dData.InternalConfigDate,
                  LightingLiveConfig: dData.LightingLiveConfig,
                  LightingLiveConfigAssignedToEmail: dData.LightingLiveConfigAssignedTo
                    ? dData.LightingLiveConfigAssignedTo.EMail
                    : '',
                  LightingLiveConfigAssignedToId: dData.LightingLiveConfigAssignedTo
                    ? dData.LightingLiveConfigAssignedTo.ID
                    : '',
                  LightingLiveConfigDate: dData.InternalConfigDate,
                  HVACConfig: dData.HVACConfig,
                  HVACConfigAssignedToEmail: dData.HVACConfigAssignedTo
                    ? dData.HVACConfigAssignedTo.EMail
                    : '',
                  HVACConfigAssignedToId: dData.HVACConfigAssignedTo
                    ? dData.HVACConfigAssignedTo.ID
                    : '',
                  HVACConfigDate: dData.HVACConfigDate,
                  OtherNetworkIpsOnsite: dData.OtherNetworkIpsOnsite,
                  OtherNetworkIpsAssignedToEmail: dData.OtherNetworkIpsAssignedTo
                    ? dData.OtherNetworkIpsAssignedTo.EMail
                    : '',
                  OtherNetworkIpsAssignedToId: dData.OtherNetworkIpsAssignedTo
                    ? dData.OtherNetworkIpsAssignedTo.ID
                    : '',
                  OtherNetworkIpsOnsiteDate: dData.OtherNetworkIpsOnsiteDate,
                  StationName: dData.StationName,
                  StationNameAssignedToEmail: dData.StationNameAssignedTo
                    ? dData.StationNameAssignedTo.EMail
                    : '',
                  StationNameAssignedToId: dData.StationNameAssignedTo
                    ? dData.StationNameAssignedTo.ID
                    : '',
                  StationNameDate: dData.StationNameDate,
                  OnboardToSupervisor: dData.OnboardToSupervisor,
                  OnboardAssignedToEmail: dData.OnboardAssignedTo
                    ? dData.OnboardAssignedTo.EMail
                    : '',
                  OnboardAssignedToId: dData.OnboardAssignedTo
                    ? dData.OnboardAssignedTo.ID
                    : '',
                  OnboardDate: dData.OnboardDate,
                  Id: dData.Id,
                  Notes: dData.Notes,
                  TaskDetails: dData.TaskDetails
                    ? JSON.parse(dData.TaskDetails)
                    : [],
                  ShippingDetails: dData.ShippingDetails
                    ? JSON.parse(dData.ShippingDetails)
                    : [],
                })
              })
              setnewtasks(InternalItems[0].TaskDetails)
              setnewShippingDetails(InternalItems[0].ShippingDetails)
              setinternalFormItem(InternalItems[0])
            }
          })
          .then(() => {
            setReRender(false)
          })
      } else {
        setReRender(false)
      }
    }
  }, [reRender])

  async function handleChange(newValue, param) {
    if (param == 'BENumber')
      setinternalFormItem({ ...internalFormItem, BENumber: newValue })
    else if (param == 'ProjectName')
      setinternalFormItem({ ...internalFormItem, ProjectName: newValue })
    else if (param == 'BuilderInitials')
      setinternalFormItem({ ...internalFormItem, BuilderInitials: newValue })
    else if (param == 'ProjectManager')
      setinternalFormItem({ ...internalFormItem, ProjectManager: newValue })
    else if (param == 'ProjectDescription')
      setinternalFormItem({
        ...internalFormItem,
        ProjectDescription: newValue,
      })
    else if (param == 'Longtitude')
      setinternalFormItem({ ...internalFormItem, Longtitude: newValue })
    else if (param == 'Latitude')
      setinternalFormItem({ ...internalFormItem, Latitude: newValue })
    else if (param == 'Address')
      setinternalFormItem({ ...internalFormItem, Address: newValue })
    else if (param == 'State')
      setinternalFormItem({ ...internalFormItem, State: newValue })
    else if (param == 'Zipcode')
      setinternalFormItem({ ...internalFormItem, Zipcode: newValue })
    else if (param == 'UtilityNetInformation')
      setinternalFormItem({
        ...internalFormItem,
        UtilityNetInformation: newValue,
      })
    else if (param == 'DateofQuote')
      setinternalFormItem({
        ...internalFormItem,
        DateofQuote: newValue,
      })
    else if (param == 'DateQuoteSent')
      setinternalFormItem({
        ...internalFormItem,
        DateQuoteSent: newValue,
      })
    else if (param == 'ShippingDate')
      setinternalFormItem({
        ...internalFormItem,
        ShippingDate: newValue,
      })
    else if (param == 'POIssued')
      setinternalFormItem({
        ...internalFormItem,
        POIssued: newValue,
      })
    else if (param == 'TrackingNumber')
      setinternalFormItem({
        ...internalFormItem,
        TrackingNumber: newValue,
      })
    else if (param == 'InternalConfigDate')
      setinternalFormItem({
        ...internalFormItem,
        InternalConfigDate: newValue,
      })
    else if (param == 'InternalConfigAssignedTo' && newValue.length > 0)
      setinternalFormItem({
        ...internalFormItem,
        InternalConfigAssignedToEmail: newValue[0].secondaryText,
        InternalConfigAssignedToId: newValue[0].id,
      })
    else if (param == 'InternalConfigAssignedTo' && newValue.length == 0)
      setinternalFormItem({
        ...internalFormItem,
        InternalConfigAssignedToEmail: '',
        InternalConfigAssignedToId: 0,
      })
    else if (param == 'InternalConfig')
      setinternalFormItem({ ...internalFormItem, InternalConfig: newValue })
    else if (param == 'LightingLiveConfig')
      setinternalFormItem({
        ...internalFormItem,
        LightingLiveConfig: newValue,
      })
    else if (param == 'LightingLiveConfigAssignedTo' && newValue.length > 0)
      setinternalFormItem({
        ...internalFormItem,
        LightingLiveConfigAssignedToEmail: newValue[0].secondaryText,
        LightingLiveConfigAssignedToId: newValue[0].id,
      })
    else if (param == 'LightingLiveConfigAssignedTo' && newValue.length == 0)
      setinternalFormItem({
        ...internalFormItem,
        LightingLiveConfigAssignedToEmail: '',
        LightingLiveConfigAssignedToId: 0,
      })
    else if (param == 'LightingLiveConfigDate')
      setinternalFormItem({
        ...internalFormItem,
        LightingLiveConfigDate: newValue,
      })
    else if (param == 'HVACConfig')
      setinternalFormItem({ ...internalFormItem, HVACConfig: newValue })
    else if (param == 'HVACConfigAssignedTo' && newValue.length > 0)
      setinternalFormItem({
        ...internalFormItem,
        HVACConfigAssignedToEmail: newValue[0].secondaryText,
        HVACConfigAssignedToId: newValue[0].id,
      })
    else if (param == 'HVACConfigAssignedTo' && newValue.length == 0)
      setinternalFormItem({
        ...internalFormItem,
        HVACConfigAssignedToEmail: '',
        HVACConfigAssignedToId: 0,
      })
    else if (param == 'HVACConfigDate')
      setinternalFormItem({ ...internalFormItem, HVACConfigDate: newValue })
    else if (param == 'OtherNetworkIpsOnsite')
      setinternalFormItem({
        ...internalFormItem,
        OtherNetworkIpsOnsite: newValue,
      })
    else if (param == 'OtherNetworkIpsAssignedTo' && newValue.length > 0)
      setinternalFormItem({
        ...internalFormItem,
        OtherNetworkIpsAssignedToEmail: newValue[0].secondaryText,
        OtherNetworkIpsAssignedToId: newValue[0].id,
      })
    else if (param == 'OtherNetworkIpsAssignedTo' && newValue.length == 0)
      setinternalFormItem({
        ...internalFormItem,
        OtherNetworkIpsAssignedToEmail: '',
        OtherNetworkIpsAssignedToId: 0,
      })
    else if (param == 'OtherNetworkIpsOnsiteDate')
      setinternalFormItem({
        ...internalFormItem,
        OtherNetworkIpsOnsiteDate: newValue,
      })
    else if (param == 'StationName')
      setinternalFormItem({ ...internalFormItem, StationName: newValue })
    else if (param == 'StationNameAssignedTo' && newValue.length > 0)
      setinternalFormItem({
        ...internalFormItem,
        StationNameAssignedToEmail: newValue[0].secondaryText,
        StationNameAssignedToId: newValue[0].id,
      })
    else if (param == 'StationNameAssignedTo' && newValue.length == 0)
      setinternalFormItem({
        ...internalFormItem,
        StationNameAssignedToEmail: '',
        StationNameAssignedToId: 0,
      })
    else if (param == 'StationNameDate')
      setinternalFormItem({ ...internalFormItem, StationNameDate: newValue })
    else if (param == 'OnboardToSupervisor')
      setinternalFormItem({
        ...internalFormItem,
        OnboardToSupervisor: newValue,
      })
    else if (param == 'OnboardAssignedTo' && newValue.length > 0)
      setinternalFormItem({
        ...internalFormItem,
        OnboardAssignedToEmail: newValue[0].secondaryText,
        OnboardAssignedToId: newValue[0].id,
      })
    else if (param == 'OnboardAssignedTo' && newValue.length == 0)
      setinternalFormItem({
        ...internalFormItem,
        OnboardAssignedToEmail: '',
        OnboardAssignedToId: 0,
      })
    else if (param == 'OnboardDate')
      setinternalFormItem({ ...internalFormItem, OnboardDate: newValue })
    else if (param == 'Notes')
      setinternalFormItem({ ...internalFormItem, Notes: newValue })
  }

  async function submitItem() {
    var viewItem = internalFormItem
    if (viewItem.Id == 0 || !viewItem.Id) {
      let list = props.spcontext.web.lists
        .getByTitle('InternalFormList')
        .items.add({
          BENumber: viewItem.BENumber.toString(),
          ProjectName: viewItem.ProjectName.toString(),
          BuilderInitials: viewItem.BuilderInitials.toString(),
          ProjectManager: viewItem.ProjectManager.toString(),
          ProjectDescription: viewItem.ProjectDescription.toString(),
          Longtitude: viewItem.Longtitude.toString(),
          Latitude: viewItem.Latitude.toString(),
          StreetAddress: viewItem.Address,
          State: viewItem.State,
          Zipcode: viewItem.Zipcode,
          DateofQuote: viewItem.DateofQuote,
          DateQuoteSent: viewItem.DateQuoteSent,
          POIssued: viewItem.POIssued,
          TrackingNumber: viewItem.TrackingNumber,
          ShippingDate: viewItem.ShippingDate,
          InternalConfig: viewItem.InternalConfig,
          InternalConfigAssignedToId: viewItem.InternalConfigAssignedToId
            ? viewItem.InternalConfigAssignedToId
            : null,
          InternalConfigDate: viewItem.InternalConfigDate,
          LightingLiveConfig: viewItem.LightingLiveConfig,
          LightingLiveConfigAssignedToId: viewItem.LightingLiveConfigAssignedToId
            ? viewItem.LightingLiveConfigAssignedToId
            : null,
          LightingLiveConfigDate: viewItem.InternalConfigDate,
          HVACConfig: viewItem.HVACConfig,
          HVACConfigAssignedToId: viewItem.HVACConfigAssignedToId
            ? viewItem.HVACConfigAssignedToId
            : null,
          HVACConfigDate: viewItem.HVACConfigDate,
          OtherNetworkIpsOnsite: viewItem.OtherNetworkIpsOnsite,
          OtherNetworkIpsAssignedToId: viewItem.OtherNetworkIpsAssignedToId
            ? viewItem.OtherNetworkIpsAssignedToId
            : null,
          OtherNetworkIpsOnsiteDate: viewItem.OtherNetworkIpsOnsiteDate,
          StationName: viewItem.StationName,
          StationNameAssignedToId: viewItem.StationNameAssignedToId
            ? viewItem.StationNameAssignedToId
            : null,
          StationNameDate: viewItem.StationNameDate,
          OnboardToSupervisor: viewItem.OnboardToSupervisor,
          OnboardAssignedToId: viewItem.OnboardAssignedToId
            ? viewItem.OnboardAssignedToId
            : null,
          OnboardDate: viewItem.OnboardDate,
          Notes: viewItem.Notes,
          TaskDetails: JSON.stringify(newtasks),
          ShippingDetails: JSON.stringify(newShippingDetails),
          ReferenceID: requestID.toString(),
          RecordType: requestType.toLocaleLowerCase() == 'wf' ? 'WF' : 'NWF',
        })
        .then(() => {
          setHideDialog(false)
        })
    } else {
      let list = props.spcontext.web.lists
        .getByTitle('InternalFormList')
        .items.getById(viewItem.Id)
        .update({
          BENumber: viewItem.BENumber,
          ProjectName: viewItem.ProjectName,
          BuilderInitials: viewItem.BuilderInitials,
          ProjectManager: viewItem.ProjectManager,
          ProjectDescription: viewItem.ProjectDescription,
          Longtitude: viewItem.Longtitude,
          Latitude: viewItem.Latitude,
          StreetAddress: viewItem.Address,
          State: viewItem.State,
          Zipcode: viewItem.Zipcode,
          DateofQuote: viewItem.DateofQuote,
          DateQuoteSent: viewItem.DateQuoteSent,
          POIssued: viewItem.POIssued,
          TrackingNumber: viewItem.TrackingNumber,
          ShippingDate: viewItem.ShippingDate,
          UtilityNetInformation: viewItem.UtilityNetInformation,
          InternalConfig: viewItem.InternalConfig,
          InternalConfigAssignedToId: viewItem.InternalConfigAssignedToId
            ? viewItem.InternalConfigAssignedToId
            : null,
          InternalConfigDate: viewItem.InternalConfigDate,
          LightingLiveConfig: viewItem.LightingLiveConfig,
          LightingLiveConfigAssignedToId: viewItem.LightingLiveConfigAssignedToId
            ? viewItem.LightingLiveConfigAssignedToId
            : null,
          LightingLiveConfigDate: viewItem.InternalConfigDate,
          HVACConfig: viewItem.HVACConfig,
          HVACConfigAssignedToId: viewItem.HVACConfigAssignedToId
            ? viewItem.HVACConfigAssignedToId
            : null,
          HVACConfigDate: viewItem.HVACConfigDate,
          OtherNetworkIpsOnsite: viewItem.OtherNetworkIpsOnsite,
          OtherNetworkIpsAssignedToId: viewItem.OtherNetworkIpsAssignedToId
            ? viewItem.OtherNetworkIpsAssignedToId
            : null,
          OtherNetworkIpsOnsiteDate: viewItem.OtherNetworkIpsOnsiteDate,
          StationName: viewItem.StationName,
          StationNameAssignedToId: viewItem.StationNameAssignedToId
            ? viewItem.StationNameAssignedToId
            : null,
          StationNameDate: viewItem.StationNameDate,
          OnboardToSupervisor: viewItem.OnboardToSupervisor,
          OnboardAssignedToId: viewItem.OnboardAssignedToId
            ? viewItem.OnboardAssignedToId
            : null,
          OnboardDate: viewItem.OnboardDate,
          Notes: viewItem.Notes,
          TaskDetails: JSON.stringify(newtasks),
          ShippingDetails: JSON.stringify(newShippingDetails),
          ReferenceID: requestID.toString(),
          RecordType: requestType.toLocaleLowerCase() == 'wf' ? 'WF' : 'NWF',
        })
        .then(() => {
          setHideDialog(false)
        })
    }
  }

  async function dynamictaskhandlechange(newValue, key, index) {
    if (key == 'TaskChecked') newtasks[index].TaskChecked = newValue

    if (key == 'TaskName') newtasks[index].TaskName = newValue

    if (key == 'CompletedBy') {
      if (newValue.length > 0) {
        newtasks[index].CompletedBy = newValue[0].secondaryText
        newtasks[index].Userid = newValue[0].id
      } else {
        newtasks[index].CompletedBy = ''
        newtasks[index].Userid = ''
      }
    }

    if (key == 'Date') newtasks[index].Date = newValue

    setnewtasks([...newtasks])
  }

  async function dynamicshippinghandlechange(newValue, key, index) {
    if (key == 'TrackingNumber')
      newShippingDetails[index].TrackingNumber = newValue

    if (key == 'ShippingDate') newShippingDetails[index].ShippingDate = newValue

    if (key == 'CarrierNumber')
      newShippingDetails[index].CarrierNumber = newValue

    setnewShippingDetails([...newShippingDetails])
  }

  async function deleteshippinghandlechange(key) {
    var deleteShippingDetails = newShippingDetails
    deleteShippingDetails.splice(key, 1)
    setnewShippingDetails([...deleteShippingDetails])
  }

  async function deletehandlechange(key) {
    var deletetasks = newtasks
    deletetasks.splice(key, 1)
    setnewtasks([...deletetasks])
  }

  var test = internalFormItem.POIssued

  return (
    <ThemeProvider
      theme={requestType.toLocaleLowerCase() == 'wf' ? redTheme : blueTheme}
    >
      <div style={{ backgroundColor: '#F2F2F2', padding: '1rem' }}>
        <div className={styles.formHeader}>
          <div>
            <Icon
              iconName="NavigateBack"
              styles={{
                root: {
                  fontSize: 30,
                  fontWeight: 600,
                  color: requestType == 'wf' ? '#d71e2b' : '#004fa2',
                },
              }}
              onClick={() =>
                (window.location.href =
                  props.context.pageContext.web.absoluteUrl +
                  `/SitePages/AdminDashboard.aspx`)
              }
            />
          </div>
          <h1 className={styles.heading}>Internal Form</h1>
          <div className={styles.SubmitSection}>
            {' '}
            <DefaultButton
              text="Devices"
              style={{ color: requestType == 'wf' ? '#d71e2b' : '#004fa2' }}
              href={
                props.context.pageContext.web.absoluteUrl +
                `/SitePages/DeviceList.aspx?RequestID=${requestID}&RequestType=${requestType}`
              }
            />
          </div>
        </div>

        <div className={styles.quoteFormSection}>
          <div
            className={styles.quoteFormSection}
            style={{ marginLeft: '0.3rem' }}
          >
            <div
              style={{
                display: 'flex',
                justifyContent: 'space-between',
                marginLeft: '0.2rem',
              }}
            >
              <TextField
                label={
                  requestType.toLocaleLowerCase() == 'wf' ? 'BE#' : 'Site Name'
                }
                styles={halfWidthInput}
                //value={internalFormItem.BENumber}
                value={internalFormItem.BENumber}
                //disabled ={true}
                onChange={(e, newValue) => handleChange(newValue, 'BENumber')}
              />
              <TextField
                label="Project Name"
                //disabled ={true}
                styles={halfWidthInput}
                value={internalFormItem.ProjectName}
                onChange={(e, newValue) =>
                  handleChange(newValue, 'ProjectName')
                }
              />
              <TextField
                label="Builder Initials"
                styles={halfWidthInput}
                value={internalFormItem.BuilderInitials}
                onChange={(e, newValue) =>
                  handleChange(newValue, 'BuilderInitials')
                }
              />
              <TextField
                label="Project Manager"
                styles={halfWidthInput}
                value={internalFormItem.ProjectManager}
                onChange={(e, newValue) =>
                  handleChange(newValue, 'ProjectManager')
                }
              />
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              <TextField
                styles={halfWidthInput}
                label="Longtitude"
                value={internalFormItem.Longtitude}
                onChange={(e, newValue) => handleChange(newValue, 'Longtitude')}
              />

              <TextField
                styles={halfWidthInput}
                label="Latitude"
                value={internalFormItem.Latitude}
                onChange={(e, newValue) => handleChange(newValue, 'Latitude')}
              />
              <TextField
                styles={halfWidthInput}
                label="State"
                value={internalFormItem.State}
                onChange={(e, newValue) => handleChange(newValue, 'State')}
              />
              <TextField
                styles={halfWidthInput}
                label="Zipcode"
                value={internalFormItem.Zipcode}
                onChange={(e, newValue) => handleChange(newValue, 'Zipcode')}
              />
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              <TextField
                label="Project Description"
                rows={2}
                multiline
                styles={halfWidthInput}
                value={internalFormItem.ProjectDescription}
                onChange={(e, newValue) =>
                  handleChange(newValue, 'ProjectDescription')
                }
              />
              <TextField
                label="Street Address"
                rows={2}
                multiline
                styles={halfWidthInput}
                value={internalFormItem.Address}
                onChange={(e, newValue) => handleChange(newValue, 'Address')}
              />
              {requestType.toLowerCase() == 'wf' ? (
                <TextField
                  label="UtilityNet Information"
                  rows={2}
                  multiline
                  styles={halfWidthInput}
                  value={internalFormItem.UtilityNetInformation}
                  onChange={(e, newValue) =>
                    handleChange(newValue, 'UtilityNetInformation')
                  }
                />
              ) : (
                <div style={{ width: 300, margin: '0 1rem 0.5rem 0' }}></div>
              )}
              <div style={{ width: 300, margin: '0 1rem 0.5rem 0' }}></div>
            </div>
            <hr></hr>
            <div style={{ display: 'flex', justifyContent: 'space-between' }}>
              <DatePicker
                styles={halfWidthInput}
                firstDayOfWeek={firstDayOfWeek}
                label="Date of Quote"
                placeholder="Select a date..."
                ariaLabel="Select a date"
                strings={defaultDatePickerStrings}
                value={
                  internalFormItem.DateofQuote
                    ? new Date(internalFormItem.DateofQuote)
                    : new Date()
                }
                onSelectDate={(date) => handleChange(date, 'DateofQuote')}
              />
              <DatePicker
                styles={halfWidthInput}
                firstDayOfWeek={firstDayOfWeek}
                label="Date Quote Sent"
                placeholder="Select a date..."
                ariaLabel="Select a date"
                strings={defaultDatePickerStrings}
                value={
                  internalFormItem.DateQuoteSent
                    ? new Date(internalFormItem.DateQuoteSent)
                    : new Date()
                }
                onSelectDate={(date) => handleChange(date, 'DateQuoteSent')}
              />
              <DatePicker
                styles={halfWidthInput}
                firstDayOfWeek={firstDayOfWeek}
                label="Target Shipping Date"
                placeholder="Select a date..."
                ariaLabel="Select a date"
                strings={defaultDatePickerStrings}
                value={
                  internalFormItem.ShippingDate
                    ? new Date(internalFormItem.ShippingDate)
                    : new Date()
                }
                onSelectDate={(date) => handleChange(date, 'ShippingDate')}
              />
              <ChoiceGroup
                styles={choiceGroupStyles}
                label="P.O. Issued"
                name="POIssued"
                defaultSelectedKey="A"
                options={options}
                onChange={(
                  ev: React.FormEvent<HTMLInputElement>,
                  option: any,
                ) => {
                  if (option.text == 'Yes') handleChange(true, 'POIssued')
                  else handleChange(false, 'POIssued')
                }}
                required={true}
              />
            </div>
            {newShippingDetails.length > 0 ? (
              <div>
                <hr></hr>
                <div className={styles.newShippingLabels}>
                  <div className={styles.newShippingLabel}>Shipping Date</div>
                  <div className={styles.newShippingLabel}>Tracking Number</div>
                  <div className={styles.newShippingLabel}>Carrier Number</div>
                  <div className={styles.newShippingLabel}></div>
                </div>
                {newShippingDetails.map((val, key) => (
                  <div className={styles.newShippingDetails}>
                    <DatePicker
                      styles={halfWidthInput}
                      firstDayOfWeek={firstDayOfWeek}
                      placeholder="Select a date..."
                      ariaLabel="Select a date"
                      strings={defaultDatePickerStrings}
                      value={new Date(val.ShippingDate)}
                      onSelectDate={(date) =>
                        dynamicshippinghandlechange(date, 'ShippingDate', key)
                      }
                    />
                    <TextField
                      styles={halfWidthInput}
                      value={val.TrackingNumber}
                      onChange={(e, newValue) =>
                        dynamicshippinghandlechange(
                          newValue,
                          'TrackingNumber',
                          key,
                        )
                      }
                    />
                    <TextField
                      styles={halfWidthInput}
                      value={val.CarrierNumber}
                      onChange={(e, newValue) =>
                        dynamicshippinghandlechange(
                          newValue,
                          'CarrierNumber',
                          key,
                        )
                      }
                    />
                    <div
                      style={{
                        width: 300,
                        margin: '0 1rem 0.5rem 0',
                        alignSelf: 'flex-end',
                        paddingBottom: '0.2rem',
                      }}
                    >
                      <IconButton
                        iconProps={{
                          iconName: 'Delete',
                          style: {
                            fontSize: 20,
                            color: requestType == 'wf' ? '#d71e2b' : '#004fa2',
                          },
                        }}
                        title="Delete"
                        data-index={key}
                        onClick={(e) => {
                          deleteshippinghandlechange(key)
                        }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            ) : (
              ''
            )}

            <div
              style={{
                display: 'flex',
                justifyContent: 'flex-end',
                margin: '1rem',
              }}
            >
              <DefaultButton
                text="Add"
                onClick={() => {
                  var shippingDetails = newShippingDetails
                  shippingDetails.push({
                    TrackingNumber: '',
                    CarrierNumber: '',
                    ShippingDate: new Date(),
                  })

                  setnewShippingDetails([...shippingDetails])
                }}
              />
            </div>
            <hr></hr>
            <Label className={styles.task}>Tasks</Label>
            {}
            {/* Tasks Satrting */}
            <div className={styles.formTasks}>
              <div style={{ display: 'flex', marginTop: '20px' }}>
                <Label style={{ width: '340px' }}>Task Name</Label>
              </div>
              <div
                className={styles.formpplpicker}
                style={{ marginTop: '20px' }}
              >
                {/* <label>Completed By</label> */}
                <Label>Completed by</Label>
              </div>

              <div
                className={styles.formDatepick}
                style={{ marginTop: '20px' }}
              >
                <Label>Date</Label>
              </div>
              <div style={{ alignSelf: 'flex-end', marginBottom: '8px' }}></div>
            </div>
            {/* Tasks */}
            <div className={styles.formTasks}>
              <Checkbox
                styles={{ root: { width: '342px' } }}
                label="Internal Configuration Completion Date"
                onChange={(e, checked) =>
                  handleChange(checked, 'InternalConfig')
                }
                checked={internalFormItem.InternalConfig}
              />
              <div className={styles.formpplpicker}>
                {/* <label>Completed By</label> */}
                <PeoplePicker
                  titleText=""
                  context={props.context}
                  personSelectionLimit={1}
                  groupName={''}
                  showtooltip={true}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  defaultSelectedUsers={[
                    internalFormItem.InternalConfigAssignedToEmail,
                  ]}
                  onChange={(e) => {
                    handleChange(e, 'InternalConfigAssignedTo')
                  }}
                  ensureUser={true}
                />
              </div>
              <div className={styles.formDatepick}>
                <DatePicker
                  styles={halfWidthInput}
                  label=""
                  firstDayOfWeek={firstDayOfWeek}
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  strings={defaultDatePickerStrings}
                  value={new Date(internalFormItem.InternalConfigDate)}
                  onSelectDate={(date) =>
                    handleChange(date, 'InternalConfigDate')
                  }
                />
              </div>
            </div>
            {requestType.toLowerCase() == 'wf' ? (
              <div className={styles.formTasks}>
                <Checkbox
                  styles={{ root: { width: '342px' } }}
                  label="Lighting Live Configuration Completion Date"
                  onChange={(e, checked) =>
                    handleChange(checked, 'LightingLiveConfig')
                  }
                  checked={internalFormItem.LightingLiveConfig}
                />
                <div className={styles.formpplpicker}>
                  {/* <label>Completed By</label> */}
                  <PeoplePicker
                    titleText=""
                    context={props.context}
                    personSelectionLimit={1}
                    groupName={''}
                    showtooltip={true}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    defaultSelectedUsers={[
                      internalFormItem.LightingLiveConfigAssignedToEmail,
                    ]}
                    onChange={(e) => {
                      handleChange(e, 'LightingLiveConfigAssignedTo')
                    }}
                    ensureUser={true}
                  />
                </div>
                <div className={styles.formDatepick}>
                  <DatePicker
                    styles={halfWidthInput}
                    label=""
                    firstDayOfWeek={firstDayOfWeek}
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    strings={defaultDatePickerStrings}
                    value={new Date(internalFormItem.LightingLiveConfigDate)}
                    onSelectDate={(date) =>
                      handleChange(date, 'LightingLiveConfigDate')
                    }
                  />
                </div>
              </div>
            ) : (
              ''
            )}
            {requestType.toLowerCase() == 'wf' ? (
              <div className={styles.formTasks}>
                <Checkbox
                  styles={{ root: { width: '342px' } }}
                  label="HVAC Live Configuration Completion Date"
                  onChange={(e, checked) => handleChange(checked, 'HVACConfig')}
                  checked={internalFormItem.HVACConfig}
                />

                <div className={styles.formpplpicker}>
                  {/* <label>Completed By</label> */}
                  <PeoplePicker
                    titleText=""
                    context={props.context}
                    personSelectionLimit={1}
                    groupName={''}
                    showtooltip={true}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    defaultSelectedUsers={[
                      internalFormItem.HVACConfigAssignedToEmail,
                    ]}
                    onChange={(e) => {
                      handleChange(e, 'HVACConfigAssignedTo')
                    }}
                    ensureUser={true}
                  />
                </div>

                <div className={styles.formDatepick}>
                  <DatePicker
                    styles={halfWidthInput}
                    label=""
                    firstDayOfWeek={firstDayOfWeek}
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    strings={defaultDatePickerStrings}
                    value={new Date(internalFormItem.HVACConfigDate)}
                    onSelectDate={(date) =>
                      handleChange(date, 'HVACConfigDate')
                    }
                  />
                </div>
              </div>
            ) : (
              ''
            )}

            <div className={styles.formTasks}>
              <Checkbox
                styles={{ root: { width: '342px' } }}
                label="Other Networks Ips onsite"
                onChange={(e, checked) =>
                  handleChange(checked, 'OtherNetworkIpsOnsite')
                }
                checked={internalFormItem.OtherNetworkIpsOnsite}
              />
              <div className={styles.formpplpicker}>
                {/* <label>Completed By</label> */}
                <PeoplePicker
                  titleText=""
                  context={props.context}
                  personSelectionLimit={1}
                  groupName={''}
                  showtooltip={true}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  defaultSelectedUsers={[
                    internalFormItem.OtherNetworkIpsAssignedToEmail,
                  ]}
                  onChange={(e) => {
                    handleChange(e, 'OtherNetworkIpsAssignedTo')
                  }}
                  ensureUser={true}
                />
              </div>

              <div className={styles.formDatepick}>
                <DatePicker
                  styles={halfWidthInput}
                  label=""
                  firstDayOfWeek={firstDayOfWeek}
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  strings={defaultDatePickerStrings}
                  value={new Date(internalFormItem.OtherNetworkIpsOnsiteDate)}
                  onSelectDate={(date) =>
                    handleChange(date, 'OtherNetworkIpsOnsiteDate')
                  }
                />
              </div>
            </div>

            <div className={styles.formTasks}>
              <Checkbox
                styles={{ root: { width: '342px' } }}
                label="Station Name"
                onChange={(e, checked) => handleChange(checked, 'StationName')}
                checked={internalFormItem.StationName}
              />
              <div className={styles.formpplpicker}>
                {/* <label>Completed By</label> */}
                <PeoplePicker
                  titleText=""
                  context={props.context}
                  personSelectionLimit={1}
                  groupName={''}
                  showtooltip={true}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  defaultSelectedUsers={[
                    internalFormItem.StationNameAssignedToEmail,
                  ]}
                  onChange={(e) => {
                    handleChange(e, 'StationNameAssignedTo')
                  }}
                  ensureUser={true}
                />
              </div>

              <div className={styles.formDatepick}>
                <DatePicker
                  styles={halfWidthInput}
                  label=""
                  firstDayOfWeek={firstDayOfWeek}
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  strings={defaultDatePickerStrings}
                  value={new Date(internalFormItem.StationNameDate)}
                  onSelectDate={(date) => handleChange(date, 'StationNameDate')}
                />
              </div>
            </div>
            {requestType.toLowerCase() == 'wf' ? (
              <div className={styles.formTasks}>
                <Checkbox
                  styles={{ root: { width: '342px' } }}
                  // styles={halfWidthInput}
                  label="Onboarded to Supervisor"
                  onChange={(e, checked) =>
                    handleChange(checked, 'OnboardToSupervisor')
                  }
                  checked={internalFormItem.OnboardToSupervisor}
                />
                <div className={styles.formpplpicker}>
                  {/* <label>Completed By</label> */}
                  <PeoplePicker
                    titleText=""
                    context={props.context}
                    personSelectionLimit={1}
                    groupName={''}
                    showtooltip={true}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    defaultSelectedUsers={[
                      internalFormItem.OnboardAssignedToEmail,
                    ]}
                    onChange={(e) => {
                      handleChange(e, 'OnboardAssignedTo')
                    }}
                    ensureUser={true}
                  />
                </div>

                <div className={styles.formDatepick}>
                  <DatePicker
                    styles={halfWidthInput}
                    label=""
                    firstDayOfWeek={firstDayOfWeek}
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    strings={defaultDatePickerStrings}
                    value={new Date(internalFormItem.OnboardDate)}
                    onSelectDate={(date) => handleChange(date, 'OnboardDate')}
                  />
                </div>
              </div>
            ) : (
              ''
            )}
            <div>
              {newtasks.length > 0
                ? newtasks.map((val, key) => (
                    <div className={styles.formTasks}>
                      <div style={{ display: 'flex', marginTop: '6px' }}>
                        <Checkbox
                          styles={{
                            root: { width: '25px', paddingTop: '6px' },
                          }}
                          label=""
                          onChange={(e, checked) =>
                            dynamictaskhandlechange(checked, 'TaskChecked', key)
                          }
                          checked={newtasks[key].TaskChecked}
                        />
                        <TextField
                          label=""
                          // styles={{
                          //   root: { width: 264, margin: "0 1rem 0.5rem 0" },
                          // }}
                          styles={halfWidthInput}
                          style={{ display: 'inline' }}
                          value={newtasks[key].TaskName}
                          onChange={(e, newValue) =>
                            dynamictaskhandlechange(newValue, 'TaskName', key)
                          }
                        />
                      </div>
                      <div className={styles.formpplpicker} style={{}}>
                        {/* <label>Completed By</label> */}
                        <PeoplePicker
                          titleText=""
                          context={props.context}
                          personSelectionLimit={1}
                          groupName={''}
                          showtooltip={true}
                          showHiddenInUI={false}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={1000}
                          defaultSelectedUsers={[newtasks[key].CompletedBy]}
                          onChange={(e) => {
                            dynamictaskhandlechange(e, 'CompletedBy', key)
                          }}
                          ensureUser={true}
                        />
                      </div>

                      <div className={styles.formDatepick} style={{}}>
                        <DatePicker
                          styles={halfWidthInput}
                          label=""
                          firstDayOfWeek={firstDayOfWeek}
                          placeholder="Select a date..."
                          ariaLabel="Select a date"
                          strings={defaultDatePickerStrings}
                          value={new Date(newtasks[key].Date)}
                          onSelectDate={(date) =>
                            dynamictaskhandlechange(date, 'Date', key)
                          }
                        />
                      </div>
                      <div
                        style={{ alignSelf: 'flex-end', marginBottom: '8px' }}
                      >
                        <IconButton
                          iconProps={{
                            iconName: 'Delete',
                            style: {
                              fontSize: 20,
                              color:
                                requestType == 'wf' ? '#d71e2b' : '#004fa2',
                            },
                          }}
                          title="Delte"
                          data-index={key}
                          onClick={(e) => {
                            deletehandlechange(key)
                          }}
                        />
                      </div>
                    </div>
                  ))
                : ''}
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
              <DefaultButton
                text="Add"
                onClick={() => {
                  count = count + 1

                  var tasksdetails = newtasks
                  tasksdetails.push({
                    TaskChecked: false,
                    TaskName: '',
                    CompletedBy: '',
                    Userid: '',
                    Date: new Date(),
                  })

                  setnewtasks([...tasksdetails])
                }}
              />
            </div>

            <div style={{ display: 'flex' }}>
              <TextField
                label="Notes"
                rows={4}
                multiline
                styles={halfWidthInput}
                value={internalFormItem.Notes}
                onChange={(e, newValue) => handleChange(newValue, 'Notes')}
              />
            </div>
            <div className={styles.devicebtn}>
              <PrimaryButton
                onClick={submitItem}
                text="Submit"
                style={{ marginRight: '0.6rem' }}
              />
              <DefaultButton
                text="Cancel"
                onClick={() =>
                  (window.location.href =
                    props.context.pageContext.web.absoluteUrl +
                    `/SitePages/AdminDashboard.aspx`)
                }
              />
            </div>
          </div>
        </div>
      </div>
      <div className={styles.SubmitSection}>
        <Dialog
          hidden={hideDialog}
          onDismiss={() => {
            setHideDialog(true)
            window.location.href =
              props.context.pageContext.web.absoluteUrl +
              `/SitePages/AdminDashboard.aspx`
          }}
          dialogContentProps={dialogContentProps}
        >
          <DialogFooter>
            <PrimaryButton
              onClick={() =>
                (window.location.href =
                  props.context.pageContext.web.absoluteUrl +
                  `/SitePages/AdminDashboard.aspx`)
              }
              text="Ok"
            />
          </DialogFooter>
        </Dialog>
      </div>
    </ThemeProvider>
  )
}
export default App
