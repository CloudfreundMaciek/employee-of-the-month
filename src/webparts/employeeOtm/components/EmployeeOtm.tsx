import { DefaultButton, Dropdown, IDropdownOption, IIconProps, Image, IStackTokens, PrimaryButton, Separator, Spinner, Stack, Text, TextField } from 'office-ui-fabric-react';
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
        const stackTokenStyles: IStackTokens = {childrenGap: '5px'};
        let t1, t2;
        if (this.state.eotm) {
          t1 = <Text><strong>{this.state.eotm.FirstName+' '+this.state.eotm.LastName}</strong> has been chosen the best employee this month!</Text>;
          t2 = <Text>{this.state.eotm.Eotm}</Text>;
        }
        else {
          t1 = <Text>The Employee of the month hasn't been chosen yet...</Text>;
          t2 = <Text>Wait for the next selection and do your best!</Text>;
        }
        currentMode = <>
        <Stack style={{height: '200px', width: 'fit-content', border: '2px solid #0078d4', borderRadius: '10px', padding: '5px'}} horizontal={true} tokens={stackTokenStyles}>
          <div style={{position: 'relative'}}>
            <Image style={{height: '185px'}} src={`${this.props.rootLink}/SiteAssets/eotm_photographs/eotm_pic.png`} />
            <Image style={{position: 'absolute', top: '11px', left: ' 32px'}} src={this.state.eotm ? this.state.eotm.PicUrl : `${this.props.rootLink}/SiteAssets/eotm_photographs/blank_user.jpg`} />
          </div>
          <Stack style={{position: 'relative', width: '185px', height: '100%'}}>
            {t1} <Separator/> {t2}
            <PrimaryButton style={{position: 'absolute', bottom: '0px'}} text='Choose the employee' onClick={()=>this.changeMode('choice')}/>
          </Stack>
        </Stack>
        </>
        break;

      case 'choice':
        currentMode = <Choice listOptions={this.listOptions} assign_eotm={this.assign_eotm} changeMode={this.changeMode}/>;
        break;
    }
    return currentMode;
  }
}
