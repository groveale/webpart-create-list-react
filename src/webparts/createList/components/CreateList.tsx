import * as React from 'react';
import styles from './CreateList.module.scss';
import { ICreateListProps } from './ICreateListProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class CreateList extends React.Component<ICreateListProps, {}> {
  public render(): React.ReactElement<ICreateListProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      uniqueWebpartId,
      doesListExist
    } = this.props;

    return (
      <section className={`${styles.createList} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            This webpart illustrates how to create and delete an associated list safely
          </p>
          <p>
            There are two buttons in the property, one to create a list and one to delete. The list will be created with the unique id of the webpart.
          </p>
          <p>
            List: {uniqueWebpartId}
          </p>
          <h3>Exists: {doesListExist.toString()}</h3>
        </div>
      </section>
    );
  }
}
