import * as React from 'react';
import styles from './AnonymusApiwpDemo.module.scss';
import type { IAnonymusApiwpDemoProps } from './IAnonymusApiwpDemoProps';

export default class AnonymusApiwpDemo extends React.Component<IAnonymusApiwpDemoProps> {
  public render(): React.ReactElement<IAnonymusApiwpDemoProps> {
    const {
      
      
      hasTeamsContext
     
    } = this.props;

    return (
      <section className={`${styles.anonymusApiwpDemo} ${hasTeamsContext ? styles.teams : ''}`}>
       <div className={styles.anonymusApiwpDemo}>
     <span className={styles.welcome}>User Details:</span>

  <div><strong>Developer Name:</strong> Hirva</div><br />

  
  <div><strong>Description:</strong> 
    React & SharePoint Developer Intern with experience in SPFx, Power Automate, and SharePoint customization.
  </div><br />

  <div><strong>Developer Email:</strong> hirva@example.com</div><br />

      <div><strong>ID:</strong>{this.props.id}</div><br />
      <div><strong>User Name:</strong>{this.props.username}</div><br />
      <div><strong>Name:</strong>{this.props.name}</div><br />
      <div><strong>Address:</strong>{this.props.address}</div><br />
      <div><strong>Email:</strong>{this.props.email}</div><br />
      <div><strong>Phone:</strong>{this.props.phone}</div><br />
      <div><strong>Website:</strong>{this.props.website}</div><br />
      <div><strong>Company:</strong>{this.props.company}</div><br />

     </div>
      </section>
    );
  }
}
