import * as React from 'react';
import styles from './Outlook2SharePoint.module.scss';
import * as strings from 'Outlook2SharePointWebPartStrings';
import GraphController from '../../../controller/GraphController';
import Groups from './Groups';
import OneDrive from './OneDrive';
import Teams from './Teams';
import { IOutlook2SharePointProps } from './IOutlook2SharePointProps';
import { IOutlook2SharePointState } from './IOutlook2SharePointState';
import { FontIcon } from 'office-ui-fabric-react/lib/Icon';
import { ILabelStyles, IStyleSet, MessageBar, MessageBarType, Pivot, PivotItem, PrimaryButton } from '@fluentui/react';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
initializeIcons("https://static2.sharepointonline.com/files/fabric/assets/icons/");

export default class Outlook2SharePoint extends React.Component<IOutlook2SharePointProps, IOutlook2SharePointState> {
  private graphController: GraphController;
  private saveMetadata = true; // For simplicity reasons and as I am not convinced with the current "Property handling" of Office Add-In we configure 'hard-coded'
  
  constructor(props) {
    super(props);
    this.state = {
      graphController: null,
      mailMetadata: null,
      showError: false,
      showSuccess: false,
      showOneDrive: false,
      showTeams: false,
      showGroups: false,
      successMessage: '',
      errorMessage: ''
    };
    this.graphController = new GraphController(this.saveMetadata);
    this.graphController.init(this.props.msGraphClientFactory)
      .then((controllerReady) => {
        if (controllerReady) {
          this.graphClientReady();
        }        
      });
  }

  public render(): React.ReactElement<IOutlook2SharePointProps> {
    const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
      root: { marginTop: 10 },
    };

    return (
      
      <div className={ styles.outlook2SharePoint }>
        {this.state.mailMetadata !== null &&
          <div className={styles.metadata}>
            <div><FontIcon iconName="InfoSolid" /> {strings.SaveInfo}</div>
            <div className={styles.subMetadata}>{strings.To} <a href={this.state.mailMetadata.saveUrl}>{this.state.mailMetadata.saveDisplayName}</a></div>
            <div className={styles.subMetadata}>{strings.On} <span>{this.state.mailMetadata.savedDate.toLocaleDateString()}</span></div>
          </div>}

        <Pivot aria-label="Basic Pivot Example">
          <PivotItem headerText="OneDrive"> 
            <br/>
            <OneDrive 
              graphController={this.state.graphController} 
              mail={this.props.mail}
              successCallback={this.showSuccess}
              errorCallback={this.showError}>
            </OneDrive>
          </PivotItem>
          <PivotItem headerText="Teams">
            <br/>
            <Teams 
              graphController={this.state.graphController} 
              mail={this.props.mail}
              successCallback={this.showSuccess}
              errorCallback={this.showError}>
            </Teams>
          </PivotItem>
          <PivotItem headerText="SharePoint">
            <br/>
            <Groups 
              graphController={this.state.graphController} 
              mail={this.props.mail}
              successCallback={this.showSuccess}
              errorCallback={this.showError}>
            </Groups>
          </PivotItem>
        </Pivot>


        
        {this.state.showSuccess && <div>
          <MessageBar
            messageBarType={MessageBarType.success}
            isMultiline={false}
            onDismiss={this.closeMessage}            
            dismissButtonAriaLabel="Close"
            truncated={true}
            overflowButtonAriaLabel="See more"
          >
            {this.state.successMessage}
          </MessageBar>
        </div>}  
        {this.state.showError && <div>
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={this.closeMessage}          
            dismissButtonAriaLabel="Close"
            truncated={true}
            overflowButtonAriaLabel="See more"
          >
            {this.state.errorMessage}
          </MessageBar>
        </div>} 
      </div>
    );
  }

  
  /**
   * This function first retrieves all OneDrive root folders from user
   */
  private graphClientReady = () => {              
    this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
      return {
        graphController: this.graphController
      };
    }); 
    if (this.saveMetadata) {
      this.getMetadata();
    }    
  }   
  
  private showError = (message: string) => {
    this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
      return {
        showError: true,
        showSuccess: false,
        errorMessage: message
      };
    });
  }

  private showSuccess = (message: string) => {
    this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
      return {
        showSuccess: true,
        showError: false,
        successMessage: message
      };
    });
  }

  private closeMessage = () => {
    this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
      return {
        showSuccess: false,
        showError: false
      };
    });
  }

  /**
   * This function expands the Teams section and collapses the other ones
   */
  private showTeams = () => {
    this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
      return {
        showTeams: true,
        showOneDrive: false,
        showGroups: false
      };
    });
  }

  /**
   * This function expands the OneDrive section and collapses the other ones
   */
  private showOneDrive = () => {
    this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
      return {
        showOneDrive: true,
        showTeams: false,
        showGroups: false
      };
    });
  }

  /**
   * This function expands the Groups section and collapses the other ones
   */
  private showGroups = () => {
    this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
      return {
        showGroups: true,
        showTeams: false,
        showOneDrive: false
      };
    });
  }

  private getMetadata() {
    this.state.graphController.retrieveMailMetadata(this.props.mail.id)
      .then((response) => {
        if (response !== null) {
          this.setState((prevState: IOutlook2SharePointState, props: IOutlook2SharePointProps) => {
            return {
              mailMetadata: response
            };
          });
        }
      });
  }
}
