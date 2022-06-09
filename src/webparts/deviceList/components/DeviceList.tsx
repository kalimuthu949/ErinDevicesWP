import * as React from 'react';
import styles from './DeviceList.module.scss';
import { IDeviceListProps } from './IDeviceListProps';
import { escape } from '@microsoft/sp-lodash-subset';
import App from "./App";

export default class DeviceList extends React.Component<IDeviceListProps, {}> {

  public render(): React.ReactElement<IDeviceListProps> {
    return (
      <App
      spcontext={this.props.spcontext} context={this.props.context}
    />
    );
  }
}
