import * as React from 'react'
import {Form, FormGroup, Button, FormControl, FormLabel, FormText} from 'react-bootstrap'
import { OutlookCategorizer } from './OutlookCategorizer';


class APIClientForm extends React.Component<{}, {token: string}> {
  constructor(props: any) {
    super(props);
    this.state = { token: '' };

    this.handleChange = this.handleChange.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
  }

  handleSubmit(event: any) {
    new OutlookCategorizer(this.state.token);
    event.preventDefault();
  }

  handleChange(event: any) {
    this.setState({ token: event.target.value });
  }

  render() {
    return (
      <Form onSubmit={this.handleSubmit}>
        <FormGroup>
        <FormLabel>
          Graph API Explorer Access Token:
         </FormLabel>
          <FormControl placeholder="Graph API Access token"  as="textarea" rows={6} type="text" size="sm" name="accesstoken" id="tokenfield" onChange={this.handleChange}/>
          <FormText>
            This API token is not stored. It is only used locally in your browser to use the Microsoft Graph API directly.
          </FormText>
        </FormGroup>
        <Button variant="primary" type="submit">Create virtual category folders</Button>
      </Form>
     
    );
  }
}

function App() {
  return (
    <APIClientForm />
  );
}

export default App;
