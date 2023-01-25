import { DefaultButton, Dropdown, IDropdownOption, IIconProps, PrimaryButton, Spinner, Stack, TextField } from 'office-ui-fabric-react';
import * as React from 'react';
import { IEmployee } from '../EmployeeOtmWebPart';
//import styles from './EmployeeOtm.module.scss';
import { IEmployeeOtmProps } from './IEmployeeOtmProps';

const icon: IIconProps = {iconName: 'SchoolDataSyncLogo'};

interface IChoiceProps {
  listOptions: ()=>Array<IDropdownOption>;
  assign_eotm: (newEotm: IEmployee, reason: string)=>Promise<void>;
  changeMode: (mode: string)=>void;
}

interface IChoiceState {
  currentEmp: IEmployee;
  loading: boolean;
}

interface IEmployeeOtmState {
  mode: string;
  eotm: IEmployee;
}

interface IProfileOption {
  key: string; 
  text: string; 
  emp: IEmployee;
}

class Choice extends React.Component<IChoiceProps, IChoiceState> {
  textRef: any = null;
  constructor(props: IChoiceProps) {
    super(props);

    this.state = {
      currentEmp: null,
      loading: false
    }
  }

  public render(): JSX.Element {
    return (
    <Stack style={{width: '250px'}} tokens={{childrenGap: '5px'}}>
      <Dropdown 
      options={this.props.listOptions()} 
      onChange={async (ev: React.FormEvent<HTMLDivElement>, option: IProfileOption) => { await this.setState({currentEmp: option.emp}); }   }
      label="Choose an employee" />

      {this.state.currentEmp && this.state.currentEmp.FirstName !== 'None' ? 
        <TextField placeholder='Reason for your decision...' componentRef={(item: any)=>{this.textRef = item;}}/> : null}
      <Stack horizontal={true} tokens={{childrenGap: '5px'}}>
        <DefaultButton text='Back' onClick={()=>this.props.changeMode('opening')} />
        {this.state.currentEmp ?
        <PrimaryButton 
        iconProps={icon} 
        text="Assign achievement" 
        onClick={()=>{
          this.setState({ loading: true });
          this.props.assign_eotm(this.state.currentEmp, this.textRef?.value)
          .then(()=>{ this.setState({ loading: false }); this.props.changeMode('opening');},
          (reason)=>console.log(reason)); 
          }} 
        />
        : null
        }
      </Stack>
      {this.state.loading ? <Spinner/> : null}
    </Stack>)
  }
}

export default class EmployeeOtm extends React.Component<IEmployeeOtmProps, IEmployeeOtmState> {

  employees: Array<IEmployee>;

  constructor(props: IEmployeeOtmProps) {
    super(props);

    this.employees = props.employees;
    
    this.state = {
      eotm: this.find_eotm(),
      mode: 'opening'
    };
    
    this.listOptions = this.listOptions.bind(this);    
    this.assign_eotm = this.assign_eotm.bind(this);
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

  public async assign_eotm(newEotm: IEmployee, reason: string = null): Promise<void> {
    return this.props.assign_eotm(this.state.eotm ? this.state.eotm : null, newEotm, reason)
    .then(async(newEotm: IEmployee | null)=>await this.setState({ eotm: newEotm }));
  }
  
  private listOptions(): Array<IProfileOption> {
    const options = new Array<IProfileOption>();

    for (const employee of this.employees) {
      options.push({
        key: employee.LoginName,
        text: employee.FirstName+' '+employee.LastName,
        emp: employee
      })
    }
    return options;
  }

  private find_eotm(): IEmployee | null {
    for (const emp of this.employees) {
      if(emp.Eotm) return emp;
    }
    return null;
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
            <p style={{margin: '0px', marginTop: '10px', fontSize: '20px'}}><strong>{this.state.eotm ? this.state.eotm.FirstName+' '+this.state.eotm.LastName : '. . .'}</strong></p>
            <p style={{fontSize: '16px', width: '80%', marginBottom: 'auto', textAlign: 'justify', hyphens: 'auto'}}>{this.state.eotm ? this.state.eotm.Eotm : "The employee of the month hasn't been chosen yet.. wait on the next selection and until then... do your best B-D"}</p>
            <PrimaryButton text='Choose the employee' onClick={()=>this.changeMode('choice')} style={{marginBottom: '16px', fontSize: '14px'}} />
          </div>
        break;

      case 'choice':
        currentMode = <Choice listOptions={this.listOptions} assign_eotm={this.assign_eotm} changeMode={this.changeMode}/>;
        break;
    }
    return currentMode;
  }
}
