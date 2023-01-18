import { ActionButton, DefaultButton, Dropdown, IDropdownOption, IIconProps, Spinner, Stack, Text, TextField } from 'office-ui-fabric-react';
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

  public render() {
    return (
    <Stack>
      <Dropdown 
      options={this.props.listOptions()} 
      onChange={async (ev: any, option: any) => { await this.setState({currentEmp: option.emp}); }   }
      label="Choose an employee" />

      {this.state.currentEmp ? 
        <TextField placeholder='Reason for your decision...' componentRef={(item: any)=>{this.textRef = item;}}/> : null}
      <Stack horizontal={true}>
        <DefaultButton text='Back' onClick={()=>this.props.changeMode('opening')} />
        {this.state.currentEmp ?
        <ActionButton 
        iconProps={icon} 
        text="Assign achievement" 
        onClick={()=>{
          this.setState({ loading: true });
          this.props.assign_eotm(this.state.currentEmp, this.textRef.value)
          .then(()=>{ this.setState({ loading: false }); this.props.changeMode('opening');}); 
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

  public async assign_eotm(newEotm: IEmployee, reason: string): Promise<void> {
    return this.props.assign_eotm(this.state.eotm ? this.state.eotm : null, newEotm, reason)
    .then(async(newEotm: IEmployee)=>await this.setState({ eotm: newEotm }));
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

  public render(): any {
    let currentMode: any;

    switch(this.state.mode) {
      case 'opening':
        currentMode = 
        <Stack>
          <Text>{this.state.eotm ? `${this.state.eotm.FirstName} ${this.state.eotm.LastName} has been chosen the best in this month. Reason: ${this.state.eotm.Eotm}` : "The employee of the month hasn't been choosed yet."}</Text>
          <ActionButton iconProps={icon} text='Set eotm' onClick={()=>this.changeMode('choice')}/>
        </Stack>;
        break;

      case 'choice':
        currentMode = <Choice listOptions={this.listOptions} assign_eotm={this.assign_eotm} changeMode={this.changeMode}/>;
        break;
    }
    return currentMode;
  }
}
