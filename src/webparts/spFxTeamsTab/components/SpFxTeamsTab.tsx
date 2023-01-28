import * as React from 'react';
import styles from './SpFxTeamsTab.module.scss';
import { ISpFxTeamsTabProps } from './ISpFxTeamsTabProps';
import { escape } from '@microsoft/sp-lodash-subset';
import * as microsoftTeams from '@microsoft/teams-js';


export default class SpFxTeamsTab extends React.Component<ISpFxTeamsTabProps, {}> {

  private teamsContext : microsoftTeams.Context ;

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve, reject) =>{
      if(this.context.microsoftTeams){
        this.context.microsoftTeams.getContext( context =>{
          this.teamsContext = context;
          resolve();
        });
      }else{
        resolve();
      }
    });
  }

  public render(): React.ReactElement<ISpFxTeamsTabProps> {

    let title : string = (this.teamsContext) 
      ? 'Teams'
      : 'SharePoint' ; 
    let currentLocation : string = (this.teamsContext)
      ? `Teams ${ this.teamsContext.teamName }`
      : `SharePoint workbench`; //${ this.context.pageContext.web.title }

    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.spFxTeamsTab} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to {title}</h3>
          <p>
            Currently in the context of following {currentLocation}
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
      </section>
    );
  }
}
