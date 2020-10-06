import * as React from 'react';
import { FontIcon } from 'office-ui-fabric-react/lib/Icon';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import styles from './Folder.module.scss';
import { IFolderProps } from './IFolderProps';

initializeIcons("https://static2.sharepointonline.com/files/fabric/assets/icons/");

export default class Folder extends React.Component<IFolderProps, {}> {  
  constructor(props) {
    super(props);
  }

  public render(): React.ReactElement<IFolderProps> {
    return (
      <li className={styles.folder}>
        <FontIcon iconName="DocLibrary" className="ms-IconDocLibrary" />&nbsp; &nbsp;                      
        <span className={`${styles.header} ${styles.isLink}`} onClick={this.getSubFolder}>{this.props.folder.name}</span>
      </li>
    );
  }

  private getSubFolder = () => {
    this.props.subFolderCallback(this.props.folder);
  }
}