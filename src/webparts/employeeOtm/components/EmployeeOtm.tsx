import { PrimaryButton } from 'office-ui-fabric-react';
import * as React from 'react';
import { IEmployee } from '../EmployeeOtmWebPart';
//import styles from './EmployeeOtm.module.scss';
import { IEmployeeOtmProps } from './IEmployeeOtmProps';

interface IEmployeeOtmState {
  mode: string;
  eotm: IEmployee;
}

export default class EmployeeOtm extends React.Component<IEmployeeOtmProps, IEmployeeOtmState> {


  constructor(props: IEmployeeOtmProps) {
    super(props);
    
    this.state = {
      eotm: this.props.Eotm,
      mode: 'opening'
    };
    
    this.assignEotm = this.assignEotm.bind(this);
    this.changeMode = this.changeMode.bind(this);
  }

  public changeMode (mode: string): void {
    switch (mode) {
      case 'opening':
        this.setState({ mode: mode });
        break;

      case 'choice':
        this.setState({ mode: mode });
        break;
    
      default:
        console.log('Uncorrect mode name!');
        break;
    }
    return;
  }

  public async assignEotm(newEotm: IEmployee, reason: string = null): Promise<void> {
    return this.props.assignEotm(this.state.eotm ? this.state.eotm : null, newEotm, reason)
    .then(async(newEotm: IEmployee | null)=>await this.setState({ eotm: newEotm }));
  }

  public render(): JSX.Element {
    let currentMode: JSX.Element;
    switch(this.state.mode) {
      case 'opening':
        currentMode = 
          <div style={{display: 'flex', flexDirection: 'column', alignItems: 'center', width: '500px', height: '400px', background: 'white', position: 'relative', border: '3px solid green', borderRadius: '10px'}}>
            <img style={{border: '2px solid gold', borderRadius: '10px', marginTop: '16px', width: '129px'}} src={this.state.eotm ? this.state.eotm.PicUrl : this.props.rootLink+'/SiteAssets/eotm_photographs/blank_user.jpg'} />
            <p style={{fontSize: '28px', fontWeight: 'bold', margin: '0px auto'}}>EMPLOYEE OF THE MONTH</p>
            <div style={{width: '80%', border: '1px solid'}}></div>
            <p style={{margin: '0px', marginTop: '10px', fontSize: '20px'}}><strong>{this.state.eotm ? this.state.eotm.Name : '. . .'}</strong></p>
            <p style={{fontSize: '16px', width: '80%', marginBottom: 'auto', textAlign: 'justify', hyphens: 'auto'}}>{this.state.eotm ? this.state.eotm.Reason : "The employee of the month hasn't been chosen yet.. wait on the next selection and until then... do your best B-D"}</p>
            <PrimaryButton text='Choose the employee' onClick={()=>this.changeMode('choice')} style={{marginBottom: '16px', fontSize: '14px'}} />
          </div>
        break;

    }
    return currentMode;
  }
}
