import * as React from 'react';
import { IEmployeeOtmProps } from './IEmployeeOtmProps';

export default class EmployeeOtm extends React.Component<IEmployeeOtmProps> {

  public render(): JSX.Element {
    return (
          <div style={{display: 'flex', flexDirection: 'column', alignItems: 'center', width: '500px', background: 'white', position: 'relative', border: '3px solid green', borderRadius: '10px', margin: '0px auto'}}>
            <img style={{border: '2px solid gold', borderRadius: '10px', marginTop: '16px', width: '129px'}} src={this.props.Eotm ? this.props.Eotm.PicUrl : this.props.rootLink+'/SiteAssets/eotm_photographs/blank_user.jpg'} />
            <p style={{margin: '0px', marginTop: '10px', fontSize: '26px'}}><strong>{this.props.Eotm ? this.props.Eotm.Name.toUpperCase() : '. . .'}</strong></p>
            <p style={{fontSize: '16px', fontWeight: 'bold', margin: '5px 0px'}}>Employee of the month</p>
            <div style={{width: '80%', border: '1px solid'}} />
            <p style={{fontSize: '16px', width: '80%', textAlign: 'center', hyphens: 'auto'}}>{this.props.Eotm ? this.props.Eotm.Reason : "The employee of the month hasn't been chosen yet.. wait on the next selection and until then... do your best B-D"}</p>
          </div>);

  }
}
