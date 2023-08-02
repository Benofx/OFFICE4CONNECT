//Le manifest est hosté ici : https://static.tpionet.com/OFFICE4CONNECT/extension-manifest.json
//Attention, en cas de message "invalid extension url" lors de la tentative d'ajout sur Connect, il est probable que cela soit dû à une erreur dans le manifest JSON.
//Cela est à controler avec un validateur de format JSON.

import logoPBI from './images/office.png';
import React from "react";
import './App.css';
import * as Extensions from "trimble-connect-workspace-api";
import isEqual from 'lodash/isEqual';
import {
  Container,
  Row,
  Col,
  ListGroup,
  ListGroupItem,
  Button,
  Modal,
  ModalBody,
  ModalFooter,
  Form,
  FormGroup,
  Input,
  Label,
  Alert,
  InputGroup,
  Badge,
} from "reactstrap";

function inIframe() {
  try {
    return window.self !== window.top;
  } catch (e) {
    return true;
  }
}

var title = ""
var link = ""

class App extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      mainMenu: { title: "Office 365", icon: "https://static.tpionet.com/OFFICE4CONNECT/office.png", command: "main_nav_menu_cliked", },
      subMenuItems: [
      { title: "Outlook", icon: `https://static.tpionet.com/OFFICE4CONNECT/outlook.png`, command: "submenu_1_clicked", },
      { title: "Sharepoint", icon: `https://static.tpionet.com/OFFICE4CONNECT/sharepoint.png`, command: "submenu_2_clicked", },
      { title: "Power BI", icon: `https://static.tpionet.com/OFFICE4CONNECT/pbi.png`, command: "submenu_2_clicked", },
      { title: "One Drive", icon: `https://static.tpionet.com/OFFICE4CONNECT/onedrive.png`, command: "submenu_2_clicked", },
      { title: "Word", icon: `https://static.tpionet.com/OFFICE4CONNECT/word.png`, command: "submenu_2_clicked", },
      { title: "Excel", icon: `https://static.tpionet.com/OFFICE4CONNECT/excel.png`, command: "submenu_2_clicked", },
      { title: "Power Point", icon: `https://static.tpionet.com/OFFICE4CONNECT/powerpoint.png`, command: "submenu_2_clicked", },
      ],
      queryParams: "?taskId=16&navigate=true",
      editParams: false,
      title: "",
      link: "",
    };
  }

  async componentDidMount() {
    const { mainMenu, subMenuItems = [] } = this.state;
    if (inIframe()) {
      this.API = await Extensions.connect(
        window.parent,
        (event, args) => {
          switch (event) {
            case "extension.command":
              switch (args.data) {
                case "main_nav_menu_cliked":
                  this.setState({ title: "Office 365" });
                  this.setState({ link: "https://www.microsoft365.com" });
                  break;
                case "submenu_1_clicked":
                  this.setState({ title: "Outlook" });
                  this.setState({ link: "https://outlook.office.com/mail" });
                  break;
                case "submenu_2_clicked":
                  this.setState({ title: "Sharepoint"});
                  this.setState({ link: "hhttps://bouyguesconstruction.sharepoint.com/_layouts/15/sharepoint.aspx" });
                break;
                case "submenu_3_clicked":
                  this.setState({ title: "Power BI"});
                  this.setState({ link: "https://app.powerbi.com/home?experience=power-bi" });
                break;
                case "submenu_4_clicked":
                  this.setState({ title: "One Drive"});
                  this.setState({ link: "https://bouyguesconstruction-my.sharepoint.com" });
                break;
                case "submenu_5_clicked":
                  this.setState({ title: "Word"});
                  this.setState({ link: "https://www.microsoft365.com/launch/word" });
                break;
                case "submenu_6_clicked":
                  this.setState({ title: "Excel"});
                  this.setState({ link: "https://www.microsoft365.com/launch/excel" });
                break;
                case "submenu_6_clicked":
                  this.setState({ title: "Power Point"});
                  this.setState({ link: "https://www.microsoft365.com/launch/powerpoint" });
                break;
              }
              break;
            case "extension.accessToken":
              this.setState({ accessToken: args.data });
              break;
            case "extension.userSettingsChanged":
              this.setState({ alertMessage: `User settings changed!` });
              this.getUserSettings();
              break;
            default:
          }
        },
        30000
      );

      this.API.ui.setMenu({ ...mainMenu, subMenus: [...subMenuItems] }).then();

      //requesting accessToken after 5 secs
      // setTimeout(() => {
      //   this.API.extension.getPermission("accesstoken").then(permission => console.log("extension:react-client:getPermission", permission));
      // }, 5000);
    }
  }

  render() {
    return (
      <iframe title={this.state.title} width="1600" height="870" src={this.state.link} frameborder="0" allowFullScreen="true">
      </iframe>
  );
  }
}

export default App;
