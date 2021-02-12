import * as React from 'react';
import styles from './Calendar.module.scss';
import { ICalendarProps } from './ICalendarProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {PageContext} from '@microsoft/sp-page-context';

export default class Calendar extends React.Component<ICalendarProps, {}> {
  public render(): React.ReactElement<ICalendarProps> {

    return (
      <>
        <h1>Web title {this.props.context.pageContext.web.title}</h1>
        <h1>Web description {this.props.context.pageContext.web.description}</h1>
        <h1>User {this.props.context.pageContext.user.displayName}</h1>
      </>
    );
  }
}
