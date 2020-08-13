import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './App';
import * as serviceWorker from './serviceWorker';
import { Container, Row, Col } from 'react-bootstrap'

ReactDOM.render(
  <React.StrictMode>
    <Container fluid='sm'>
      <Row className="justify-content-md-center">
        <Col />
        <Col xs={10}>
          <h1>Outlook virtually categorized</h1>
          <h5>This app creates cross-device, cross-client virtual
            folders for each of your Outlook categories.</h5>
          <hr />
          <p>
            For more information and the source code of this app, please visit <a href="https://github.com/berndverst/outlook-categorized">berndverst/outlook-categorized</a>.
            </p>
          <p>
            <strong>Instructions</strong>
            <ol>
              <li>
                Visit <a href="https://developer.microsoft.com/graph/graph-explorer?wt.mc_id=outlookcategorized-github-beverst">Microsoft Graph Explorer</a>&nbsp;
            to obtain an API access token for your account.
            <ol>
                  <li>Sign into your Office365 or Outlook account.</li>
                  <li>Click on the Access Token tab.</li>
                  <li>Click on the Copy icon or manually copy the token text.</li>
                </ol>
              </li>
              <li>
                Paste the access token below and press the button to automatically create new virtual folders that will sync across your devices and email clients.
          </li>
            </ol>
          </p>
          <App />
        </Col>
        <Col />
      </Row>
      &nbsp;
      <Row />
      <Row>
        <Col />
        <Col xs={10}>
          <div id='results' />
        </Col>
        <Col />
      </Row>
    </Container>
  </React.StrictMode>,
  document.getElementById('root')
);

// If you want your app to work offline and load faster, you can change
// unregister() to register() below. Note this comes with some pitfalls.
// Learn more about service workers: https://bit.ly/CRA-PWA
serviceWorker.unregister();
