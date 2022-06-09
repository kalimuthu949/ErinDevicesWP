import * as React from 'react';
import styles from './InternalForm.module.scss';
import { IInternalFormProps } from './IInternalFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import App from "./App";

export default class InternalForm extends React.Component<IInternalFormProps, {}> {

  public render(): React.ReactElement<IInternalFormProps> {
    return (
      <App
      spcontext={this.props.spcontext} context={this.props.context}
    />
    );
  }
}
