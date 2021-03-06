import * as React from 'react';
import styles from './StateLicense.module.scss';
import { IStateLicenseProps } from './IStateLicenseProps';
import { IStateLicenseState } from './IStateLicenseState';
import { escape } from '@microsoft/sp-lodash-subset';
import axios from 'axios';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';

export default class StateLicense extends React.Component<IStateLicenseProps, IStateLicenseState> {

  constructor(props) {
    super(props);
    this.submit = this.submit.bind(this);
    this.state = {
      state: undefined,
      userEmail: undefined,
      submitStatus: false
    }
  }

  public render(): React.ReactElement<IStateLicenseProps> {
    return (
      <div>
        <div className={styles.dropdown} style={{width: 67 + '%'}}>
          <Dropdown
          disabled={this.state.submitStatus}
          placeHolder="Select a State"
          label="States Available for Licensing"
          id="stateSelect"
          options={this.USStates()}
          onChanged={(option: IDropdownOption, index: any) => {this.setState({state: option.text})}}
          />
          </div>
          <div className = { styles.spacer } style = {{marginTop: 10 + 'px'}}>
          <DefaultButton
          disabled={this.state.submitStatus}
          primary={true}
          text="Submit"
          onClick={this.submit}
          />
          { this.state.submitStatus ?
          <div>
            <div className={styles.spacer} style = {{marginTop: 10 + 'px'}}> 
              <label>Thank you! Your request has been sent. You'll receive an email confirmation shortly.</label>
            </div>
            <div style = {{marginTop: 10 + 'px'}}>
              <DefaultButton
              primary={true}
              text="Submit Another Request"
              onClick={() => {this.setState({ submitStatus: false})}}
              />
            </div>
          </div>
          : null }
        </div>
      </div>
    )
  }
  private submit() {
    let digest:any = "";
    let state: string = this.state.state;
    let userEmail: string = this.state.userEmail;
    axios.post('https://peoplesmortgagecompany.sharepoint.com/sites/intranet/requestforms/_api/contextinfo')
    .then((res) => {
      this.setState({
        submitStatus: true
      })
      digest = res.data.FormDigestValue;
    })
    .then(() => {
      axios({
        method: 'POST',
        url: "https://peoplesmortgagecompany.sharepoint.com/sites/intranet/requestforms/_api/web/lists/GetByTitle('License%20Requests')/items",
        headers: {
          "X-RequestDigest": digest,
          "Accept": "application/json;odata=verbose",
          "content-type": "application/json;odata=verbose",
        },
        data: {
          '__metadata': {
            'type': 'SP.Data.License_x0020_RequestsListItem'
          },
          'Title': new Date(),
          'State': state,
          'Loan_x0020_Officer': userEmail
        }
      })
    })
    .then((res:any ) => {
      if(res.status == 200) {
        this.setState
      }
    })
  }

  private USStates() {
    let statesArr = this.props.states;
    let staterinos = statesArr.split(',');
    let options = [];
    for (let i = 0; i < staterinos.length; i++) {
      options[i] = { key: i.toString(), text: staterinos[i]}
    }
    return options;
  }

  componentDidMount():void {
    axios({
      method:'GET',
      url:'https://peoplesmortgagecompany.sharepoint.com/_api/web/CurrentUser'
    })
    .then((res) => {
      this.setState({
        userEmail: res.data.Email
      });
    });
  }
}
