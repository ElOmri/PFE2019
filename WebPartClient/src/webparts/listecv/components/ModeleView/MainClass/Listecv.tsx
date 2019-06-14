
import * as React from "react";
import styles from "../../Listecv.module.scss";
import { IListecvProps } from "../../Modele/IListecvProps";
import { CommandBarButton } from "office-ui-fabric-react/lib/Button";
import * as $ from "jquery";
import jsPDF from "jsPDF";
import Autocomplete from "react-autocomplete";
import TimeKeeper from 'react-timekeeper';
import html2canvas from "html2canvas";
import {
  IOverflowSetItemProps,
  OverflowSet
} from "office-ui-fabric-react/lib/OverflowSet";
import Slider from "react-slick";
import { Template2 } from "../Template2/template2";
import "slick-carousel/slick/slick.css";
import "slick-carousel/slick/slick-theme.css";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import {
  ComboBox,
  Fabric,
  IComboBoxOption,
  Toggle
} from "office-ui-fabric-react/lib/index";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { MessageBarButton } from "office-ui-fabric-react/lib/Button";
import { Link } from "office-ui-fabric-react/lib/Link";
import "react-intl-tel-input/dist/main.css";
import ToggleButtonGroup from "@material-ui/lab/ToggleButtonGroup";
import ToggleButton from "@material-ui/lab/ToggleButton";
import StarRateIcon from "@material-ui/icons/StarRate";
import SupervisorAccountIcon from "@material-ui/icons/SupervisorAccount";
import {
  DefaultButton,
  PrimaryButton
} from "office-ui-fabric-react/lib/Button";
import {
  MessageBar,
  MessageBarType
} from "office-ui-fabric-react/lib/MessageBar";
import { createStyles, Theme, WithStyles } from "@material-ui/core/styles";
import DayPicker from "react-day-picker";
import {
  Dialog,
  DialogType,
  DialogFooter
} from "office-ui-fabric-react/lib/Dialog";
import "react-day-picker/lib/style.css";
import { ITextFieldProps } from "office-ui-fabric-react/lib/TextField";
import { IconButton } from "office-ui-fabric-react/lib/Button";
import { Callout } from "office-ui-fabric-react/lib/Callout";
import SimpleCrypto from "simple-crypto-js";
import { getId } from "office-ui-fabric-react/lib/Utilities";
import ReactPhoneInput from "react-phone-input-2";
import "react-phone-input-2/dist/style.css";
import { IState } from "../../Modele/IState";
import { ComboBoxOptions } from "../../DataProvider/ComboBoxOptions";
import { SectionClassName } from "../../DataProvider/SectionClassName";

const stylesx = (theme: Theme) =>
  createStyles({
    margin: {
      margin: theme.spacing.unit
    }
  });
export interface Propsx extends WithStyles<typeof stylesx> {}



export default class Listecv extends React.Component<IListecvProps, IState> {
  constructor(props: IListecvProps) {
    super(props);
    this.handleTimeChange = this.handleTimeChange.bind(this)
    if (typeof Storage !== "undefined") {
      let storedNames = [];
      let storedIds = [];
      if (JSON.parse(localStorage.getItem("names")) == undefined) {
        localStorage.setItem("names", JSON.stringify(storedNames));
      }
      if (JSON.parse(localStorage.getItem("listeid")) == undefined) {
        localStorage.setItem("listeid", JSON.stringify(storedIds));
      } else {
      }
    }
    this.state = {
       besoin: "",
       typebesoin: "",
      userPrincipalNameClicked: "",
      clockTime:
      {
        hour: 10,
        minute: 0
      },
      passwordChange:false,
      id: null,
      showPanel3: false,
      images: null,
      projects: null,
      filtertext: "",
      responses2: null,
      filter: "ID",
      loggedinUser: "",
      signRequest: null,
      login: "",
      pw: "",
      hideDialog: false,
      hideDialog2: true,
      titleRequest: ["default"],
      isCalloutVisible: false,
      adresse: "",
      company: "",
      nom: "",
      prenom: "",
      identifiant: "",
      password: "",
      password2: "",
      telephone: "",
      showPanel2: false,
      userevent: null,
      showPanel: false,
      isTeachingBubbleVisible: false,
      alignment: "left",
      responses: [],
      mode: false,
      isFavorite: false,
      token: null,
      modeleTemplate: null,
      favoriID: null,
      favoriUsersPrincipalName: [],
      autoComplete: false,
      allowFreeform: true,
      selectedDay: null,
      isEmpty: true,
      isDisabled: false,
      selectedDays: []
    };
  }
  public Image;
  public _image: HTMLElement;
  private _descriptionId: string = getId("description");
  private _iconButtonId: string = getId("iconButton");

  private _onRenderItem(item: IOverflowSetItemProps): JSX.Element {
    if (item.onRender) {
      return item.onRender(item);
    }
    return (
      <CommandBarButton
        iconProps={{ iconName: item.icon }}
        menuProps={item.subMenuProps}
        text={item.name}
      />
    );
  }

  private _onRenderOverflowButton(
    overflowItems: any[] | undefined
  ): JSX.Element {
    return (
      <CommandBarButton
        menuIconProps={{ iconName: "More" }}
        menuProps={{ items: overflowItems! }}
      />
    );
  }

  private addUser(
    _componentContext,
    token,
    identifiant,
    companyName,
    password,
    adresse,
    telephone,
    nom,
    prenom
  ) {
    var _secretKey = "elomri";
    var simpleCrypto = new SimpleCrypto(_secretKey);
    var cryptedpassword = simpleCrypto.encrypt(password);
    var x = {
      fields: {
        Title: identifiant,
        company: companyName,
        password: cryptedpassword,
        adresse: adresse,
        telephone: telephone,
        nom: nom,
        email: prenom,
        verifie: "non"
      }
    };
    $.ajax({
      async: true,
      crossDomain: true,
      url:
        "https://graph.microsoft.com/beta/sites/root/lists/cec630c7-c1f1-4025-a8b2-d77167035e5d/items/",
      method: "POST",
      headers: {
        "content-type": "application/json"
      },
      data: JSON.stringify(x),
      beforeSend: function(xhr) {
        xhr.setRequestHeader("Authorization", "Bearer " + token);
      },
      success: function() {}
    });
  }

  private _onRenderLabel = (props: ITextFieldProps): JSX.Element => {
    return (
      <>
        <span>{props.label}</span>
        <IconButton
          id={this._iconButtonId}
          iconProps={{ iconName: "Info" }}
          title="Info"
          ariaLabel="Info"
          onClick={this._onIconClick}
          styles={{ root: { marginBottom: -3 } }}
        />

        {this.state.isCalloutVisible && (
          <Callout
            target={"#" + this._iconButtonId}
            setInitialFocus={true}
            onDismiss={this._onDismiss2}
            ariaDescribedBy={this._descriptionId}
            role="alertdialog"
          >
            <span id={this._descriptionId}>Jamais partager l'identifiant</span>
            <DefaultButton onClick={this._onDismiss2}>Close</DefaultButton>
          </Callout>
        )}
      </>
    );
  };

  private _onIconClick = (): void => {
    this.setState({ isCalloutVisible: !this.state.isCalloutVisible });
  };

  private _onDismiss2 = (): void => {
    this.setState({ isCalloutVisible: false });
  };

  handleDayClick(day, selected) {
    const { selectedDays } = this.state;
      if(selected.disabled)
      {
        alert("selectionner une date valide.")
      }else
      {
  
        if (selected.selected) {
          const selectedIndex = selectedDays.findIndex(
            selectedDay => selectedDay.getTime() == day.getTime()
          );
          selectedDays.splice(selectedIndex, 1);
        } else {
          let date = day.setHours(this.state.clockTime.hour,this.state.clockTime.minute)
          let result =new Date(date)
          selectedDays.push(result);
          this.setState({
            clockTime : {hour : 12, minute : 0}
          })
          alert("Date : "+result.toLocaleDateString()+" est selectionné.")
        }
        this.setState({ selectedDays  });
        console.log(this.state.selectedDays)
    
      }
    

  }

  private requestToken(_ComponenetContext): void {
    $.ajax({
      async: true,
      crossDomain: true,
      url:
        "https://cors-anywhere.herokuapp.com/https://login.microsoftonline.com/dc21cc0b-fd43-41ea-b03b-75e4fcb3aec9/oauth2/v2.0/token", 
      method: "POST",
      headers: {
        "content-type": "application/x-www-form-urlencoded"
      },
      data: {
        grant_type: "client_credentials",
        "client_id ": "6fc6e6f6-4bd1-4392-991d-e7aae8bbb357",
        client_secret: "fikB96{^adcdBNOUYA325@^",
        "scope ": "https://graph.microsoft.com/.default"
      },

      success: function(response) {
        _ComponenetContext.setState({
          token: response.access_token
        });
      }
      
    });
  }

  private _getErrorMessage2 = (value: string): string => {
    let x = this.state.titleRequest;
    x.value.forEach(element => {
      if (element.fields.Title == value) {
        this.setState({
          identifiant: ""
        });
        alert("Utilisateur utilisé");
        return `identifiant utilisé`;
      }
    });

    return "";
  };

  private Deconnect() {
    this.setState({
      loggedinUser: "",
      adresse: "",
      company: "",
      telephone: "",
      nom: "",
      prenom: ""
    });
  }

  private _getErrorMessage = (value: string): string => {
    if (value != this.state.password) {
      this.setState({
        password2: ""
      });
    }
    return value == this.state.password
      ? ""
      : `les deux mot de passe ne sont pas identiques.`;
  };

  private getCurrentUser(componentContext, token, mode): boolean {
    if (mode) {
      if (componentContext.state.alignment == "left") {
        componentContext.setState({
          alignment: "right"
        });
      } else {
        componentContext.setState({
          alignment: "left"
        });
      }
      this.setState({
        mode: !this.state.mode
      });
    }

    let verif = false;
    componentContext.state.favoriUsersPrincipalName.forEach(element => {
      var xhr = new XMLHttpRequest();
      xhr.open(
        "GET",
        "https://graph.microsoft.com/beta/users/" +
          element +
          "/?$select=id,displayName,mail,userPrincipalName,businessPhones,city,country,officeLocation,streetAddress,skills,schools, aboutMe,interests,pastProjects,companyName,department,jobTitle,responsibilities",
        true
      );
      xhr.setRequestHeader("Authorization", "Bearer " + token);

      let Response = null;
      xhr.onreadystatechange = function() {
        if (xhr.readyState === 4 && xhr.status === 200) {
          Response = JSON.parse(xhr.responseText);

          let x = componentContext.state.responses;

          if (!x.includes(Response)) {
            x.push(Response);

            componentContext.setState({
              responses: x,
              responses2: x
            });
          }

          verif = true;
        }
      };
      xhr.send();
    });

    var xhr5 = new XMLHttpRequest();
    xhr5.open(
      "GET",
      "https://graph.microsoft.com/beta/sites/root/lists/94fa4f56-2957-4d33-a71d-c8c8488469cf/items/?expand=fields",
      true
    );
    xhr5.setRequestHeader("Authorization", "Bearer " + token);

    let Response = null;
    xhr5.onreadystatechange = function() {
      if (xhr5.readyState === 4 && xhr5.status === 200) {
        Response = JSON.parse(xhr5.responseText);

        componentContext.setState({
          projects: Response
        });

        verif = true;
      }
    };
    xhr5.send();

    return verif;
  }

  private htmltoimage() {
    this.state.responses.forEach(element => {
      var node = document.getElementById(element.id);

      html2canvas(node).then(function(canvas) {
        let data = canvas.toDataURL("image/png");
        if (this.state.images) {
          let x = this.state.images;
          if (!x.includes(data)) {
            x.push(data);
            this.setState({
              images: x
            });
          }
        } else {
          let x = [];

          x.push(data);
          this.setState({
            images: x
          });
        }
      });
    });

    console.log(this.state.images);
  }

  private getlisteID(componentContext, token): void {
    var xhr = new XMLHttpRequest();
    xhr.open(
      "GET",
      "https://graph.microsoft.com/beta/sites/root/lists/227d22af-04e8-44ea-9acc-5e88a5652db3/items/?$select=id",
      true
    );
    xhr.setRequestHeader("Authorization", "Bearer " + token);

    let Response = null;
    xhr.onreadystatechange = function() {
      if (xhr.readyState === 4 && xhr.status === 200) {
        Response = JSON.parse(xhr.responseText);

        Response.value.forEach(element => {
          {
            var xhr2 = new XMLHttpRequest();

            xhr2.open(
              "GET",
              "https://graph.microsoft.com/beta/sites/root/lists/FavoriUsers/items/" +
                element.id +
                "/fields/?$select=Title",
              true
            );
            xhr2.setRequestHeader("Authorization", "Bearer " + token);

            let Response2 = null;
            xhr2.onreadystatechange = function() {
              if (xhr2.readyState === 4 && xhr2.status === 200) {
                Response2 = JSON.parse(xhr2.responseText);
                let x = componentContext.state.favoriUsersPrincipalName;

                if (!x.includes(Response2.Title)) {
                  x.push(Response2.Title);

                  componentContext.setState({
                    favoriUsersPrincipalName: x
                  });
                }
              }
            };
            xhr2.send();
          }
        });
      }
    };
    xhr.send();
  }

  private getidentifiant(token, componentContext) {
    var xhr = new XMLHttpRequest();
    xhr.open(
      "GET",
      "https://graph.microsoft.com/beta/sites/root/lists/cec630c7-c1f1-4025-a8b2-d77167035e5d/items?expand=fields(select=Title)",
      true
    );
    xhr.setRequestHeader("Authorization", "Bearer " + token);
    let Response = null;
    xhr.onreadystatechange = function() {
      if (xhr.readyState === 4 && xhr.status === 200) {
        Response = JSON.parse(xhr.responseText);

        componentContext.setState({
          titleRequest: Response
        });
      }
    };
    xhr.send();
  }
  private getevents(element, token, componentContext) {
    var xhr = new XMLHttpRequest();
    xhr.open(
      "GET",
      "https://graph.microsoft.com/beta/users/" +
        element +
        "/events?$select=subject,start",
      true
    );
    xhr.setRequestHeader("Authorization", "Bearer " + token);
    let Response = null;
    xhr.onreadystatechange = function() {
      if (xhr.readyState === 4 && xhr.status === 200) {
        Response = JSON.parse(xhr.responseText);

        componentContext.setState({
          userevent: Response
        });
      }
    };
    xhr.send();
  }

  private Slide() {
    $("#slider").empty();
    let x = [];
    if (this.state.mode) {
      const filteredArr = this.state.responses.reduce((acc, current) => {
        const x = acc.find(item => item.displayName === current.displayName);
        if (!x) {
          return acc.concat([current]);
        } else {
          return acc;
        }
      }, []);
      if (this.state.responses.length > 0) {
        let slide = 0;
        filteredArr.forEach(element => {
          slide++;
          x.push([
            <div id={"slide" + slide}>
              <div
                ref={() => (this.Image = Image)}
                id="divToPrint"
                className="mt4"
                style={{
                  width: "auto",
                  minHeight: "auto",
                  marginLeft: "auto",
                  marginRight: "auto"
                }}
              >
                <div>
                  <Template2
                    projects={this.state.projects}
                    templatemodele=""
                    id={element.id}
                    displayName={element.displayName}
                    mail={element.mail}
                    userPrincipalName={element.userPrincipalName}
                    businessPhones={element.businessPhones}
                    city={element.city}
                    country={element.country}
                    officeLocation={element.officeLocation}
                    streetAddress={element.streetAddress}
                    skills={element.skills}
                    schools={element.schools}
                    aboutMe={element.aboutMe}
                    interests={element.interests}
                    pastProjects={element.pastProjects}
                    companyName={element.companyName}
                    department={element.department}
                    jobTitle={element.jobTitle}
                    responsibilities={element.responsibilities}
                  />
                </div>
                <DefaultButton
                  text="Export PDF"
                  onClick={() => this.printDocument(element.id)}
                />
                <PrimaryButton
                  text="Contact"
                  onClick={() =>
                    this._showPanel(
                      element.userPrincipalName,
                      this.state.token,
                      this
                    )
                  }
                />

                <div>
                  <img
                    src={
                      !JSON.parse(localStorage.getItem("listeid")).includes(
                        element.userPrincipalName
                      )
                        ? "https://image.flaticon.com/icons/svg/149/149220.svg"
                        : "https://image.flaticon.com/icons/svg/148/148839.svg"
                    }
                    onClick={() => {
                      this.setFavourite(element.userPrincipalName);
                    }}
                    width={50}
                    height={50}
                    style={{ float: "right" }}
                  />
                </div>
              </div>
            </div>
          ]);
        });

        return x;
      }
    } else {
      let favouri = JSON.parse(localStorage.getItem("listeid"));
      const filteredArr = this.state.responses.reduce((acc, current) => {
        const x = acc.find(item => item.displayName === current.displayName);
        if (!x) {
          return acc.concat([current]);
        } else {
          return acc;
        }
      }, []);
      if (this.state.responses.length > 0) {
        filteredArr.forEach(element => {
          if (favouri.includes(element.userPrincipalName)) {
            x.push([
              <div>
                <div
                  ref={() => (this.Image = Image)}
                  id="divToPrint"
                  className="mt4"
                  style={{
                    width: "auto",
                    minHeight: "auto",
                    marginLeft: "auto",
                    marginRight: "auto"
                  }}
                >
                  <div>
                    <Template2
                      projects={this.state.projects}
                      id={element.id}
                      templatemodele=""
                      displayName={element.displayName}
                      mail={element.mail}
                      userPrincipalName={element.userPrincipalName}
                      businessPhones={element.businessPhones}
                      city={element.city}
                      country={element.country}
                      officeLocation={element.officeLocation}
                      streetAddress={element.streetAddress}
                      skills={element.skills}
                      schools={element.schools}
                      aboutMe={element.aboutMe}
                      interests={element.interests}
                      pastProjects={element.pastProjects}
                      companyName={element.companyName}
                      department={element.department}
                      jobTitle={element.jobTitle}
                      responsibilities={element.responsibilities}
                    />
                  </div>
                  <DefaultButton
                    text="Export PDF"
                    onClick={() => this.printDocument(element.id)}
                  />
                  <PrimaryButton
                    text="Contact"
                    onClick={() =>
                      this._showPanel(
                        element.userPrincipalName,
                        this.state.token,
                        this
                      )
                    }
                  />
                </div>
              </div>
            ]);
          }
        });

        return x;
      }
    }
  }

  private setFavourite(namex): void {
    let x = JSON.parse(localStorage.getItem("listeid"));
    if (!x.includes(namex)) {
      x.push(namex);
      localStorage.setItem("listeid", JSON.stringify(x));
    } else {
      let num = x.indexOf(namex);
      x.splice(num, 1);
      localStorage.setItem("listeid", JSON.stringify(x));
    }
    this._showDialog2();
    setTimeout(function() {
      this.setState({ hideDialog2: true });
    }, 1000);
  }

  private requestSignin(componentContext, token) {
    if (componentContext.state.signRequest == null) {
      var xhr = new XMLHttpRequest();
      xhr.open(
        "GET",
        "https://graph.microsoft.com/beta/sites/root/lists/cec630c7-c1f1-4025-a8b2-d77167035e5d/items?expand=fields",
        true
      );
      xhr.setRequestHeader("Authorization", "Bearer " + token);
      let Response = null;
      xhr.onreadystatechange = function() {
        if (xhr.readyState === 4 && xhr.status === 200) {
          Response = JSON.parse(xhr.responseText);

          componentContext.setState({
            signRequest: Response
          });
        }
      };
      xhr.send();
    }
  }

  private signin(componentContext) {
    let login = this.state.login;
    let password = this.state.pw;
    let verif = true;
    this.state.signRequest.value.forEach(element => {
      if (element.fields.Title == login) {
        let password2 = element.fields.password;
        var _secretKey = "elomri";
        var simpleCrypto = new SimpleCrypto(_secretKey);
        var decryptedPassword = simpleCrypto.decrypt(password2);
        if (element.fields.verifie == "non") {
          alert("Veuillez vérifier votre adresse mail");
          verif = false;
        } else if (password == decryptedPassword) {
          componentContext.setState({
            loggedinUser: login,
            id : element.fields.id,
            adresse: element.fields.adresse,
            company: element.fields.company,
            prenom: element.fields.email,
            nom: element.fields.nom,
            password: decryptedPassword,
            telephone: element.fields.telephone,

            hideDialog: true
          });
          verif = false;
        }
      }
    });
    if (verif) {
      alert("verifier le mot de passe ou login");
    }
  }

  private _onFilter(text: string): void {
    this.htmltoimage();
    if (text == "") {
      this.getCurrentUser(this, this.state.token, false);
    }

    switch (this.state.filter) {
      case "ID": {
        this.setState({
          responses: this.state.responses.filter(i => {
            if (i.id != null)
              return i.id.toLowerCase().indexOf(text.toLowerCase()) > -1;
          })
        });

        break;
      }
      case "Formations": {
        this.setState({
          responses: this.state.responses.filter(i => {
            if (i.schools != null) return i.schools.includes(text);
          })
        });

        break;
      }

      case "Skills": {
        this.setState({
          responses: this.state.responses.filter(i => {
            if (i.skills != null) return i.skills.includes(text);
          })
        });

        break;
      }

      case "Expériences": {
        this.setState({
          responses: this.state.responses.filter(i => {
            if (i.pastProjects != null) return i.pastProjects.includes(text);
          })
        });

        break;
      }

      case "Loisirs": {
        this.setState({
          responses: this.state.responses.filter(i => {
            if (i.interests != null) return i.interests.includes(text);
          })
        });
        break;
      }

      default:
        break;
    }
  }

  private printDocument(id) {
    const input = document.getElementById(id);
    html2canvas(input).then(canvas => {
      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF("p", "mm", "a2");
      pdf.addImage(imgData, "JPEG", 0, 0);
      pdf.save(id + ".pdf");
    });
  }

  public componentDidMount(): void {
    this.requestToken(this);
  }
  private element(): any {
    if (this.state.token) {
      return this.state.token;
    } else return "";
  }
  public componentWillUpdate(): void {}

  private _showPanel(element, token, componentContext) {
    if (this.state.loggedinUser == "") {
      alert("Veuillez connecter ");
    } else {
      this.getevents(element, token, componentContext);
      this.setState({ showPanel: true,
      userPrincipalNameClicked : element
      });
    }
  }

  private _showPanel2(componentContext, token) {
    if (this.state.loggedinUser == "") {
      this.getidentifiant(token, componentContext);
      this.setState({ showPanel2: true });
    } else {
      alert("veuillez deconnecter");
    }
  }

  private _showPanel3() {
  
    
      this.setState({ showPanel3: true });
    
  }

  private getCalender(): Date[] {
    let x = [];
    if (this.state.userevent != null) {
      this.state.userevent.value.forEach(element => {
        x.push(new Date(element.start.dateTime));
      });

      for (let index = 1; index < new Date().getDate(); index++) {
        x.push(
          new Date(new Date().getFullYear(), new Date().getMonth(), index)
        );
      }

      return x;
    } else {
      let x = [];
      x.push(new Date());
      return x;
    }
  }

  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  };

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  };
  private _showDialog2 = (): void => {
    this.setState({ hideDialog2: false });
  };

  private _closeDialog2 = (): void => {
    this.setState({ hideDialog2: true });
  };
  private _hidePanel3 = (): void => {
    this.setState({ showPanel3: false });
  };

  private _hidePanel2 = (): void => {
    this.setState({ showPanel2: false });
  };
  private _hidePanel = (): void => {
    this.setState({ showPanel: false });
  };
  public render(): React.ReactElement<IListecvProps> {
    {
    }
    const disabled = this.getCalender();

    const _ComponenetContext = this;
    const state = this.state;
    var settings = {
      dots: true,
      adaptiveHeight: false,
      width: 1200,
      speed: 700,
      slidesToShow: 1,
      slidesToScroll: 1,

      beforeChange: (m, x, current, next) =>
        setTimeout(
          () =>
            this.setState(prevState => ({ ...prevState, currentSlide: next })),
          700
        )
    };

    {
      this.getlisteID(_ComponenetContext, _ComponenetContext.state.token);
    }
    {
      this.requestSignin(_ComponenetContext, _ComponenetContext.state.token);
    }
    if (this.element() == "") {
      return <div>Loading ..</div>;
    }
    return (
      <div>
        <div>
          <Dialog
            hidden={this.state.hideDialog2}
            onDismiss={this._closeDialog2}
            dialogContentProps={{
              type: DialogType.largeHeader,
              title: "Favoris",
              subText: "les profils favoris sont sauvegardés au navigateur"
            }}
            containerClassName={"ms-dialogMainOverride " + styles.textDialog}
          >
            <div style={{ width: 230 }}>
              <img src="https://media.giphy.com/media/6zHB86JLnQHFS/giphy.gif" />
            </div>
          </Dialog>
          <div style={{ float: "right", paddingTop: 50 }}>
            <PrimaryButton
              text="Créer un compte"
              onClick={() => this._showPanel2(this, this.state.token)}
              style={{ paddingRight: 20, paddingLeft: 20 }}
            />
            <span> </span>
            <PrimaryButton
              text="Connexion"
              onClick={() => this._showDialog()}
              style={{ paddingRight: 20, paddingLeft: 20 }}
            />
          </div>
          {this.state.loggedinUser != "" ? (
            <div style={{ float: "left" }}>
              <span style={{ fontWeight: "bold" }}> Bonjour </span>
              <span style={{ fontWeight: "bold", color: "blue" }}>
                {" "}
                {this.state.loggedinUser}
              </span>
              <br />
               <PrimaryButton
                text={"Modifier le profil"}
                onClick={() => this._showPanel3()}
              />
              <DefaultButton
                text={"Deconnexion"}
                onClick={() => this.Deconnect()}
              />
             
            </div>
          ) : (
            <p />
          )}
         
<section style={{float:'left', marginTop:20 , marginBottom:20 }}  >
          <OverflowSet
          items={[
            {
              key: "search",
              onRender: () => {
                return (
                  <div>
                   
                    <Autocomplete          
  getItemValue={(item) => item.label}
  items={this.getSuggeestions()}
  renderItem={(item, isHighlighted) =>
    <div style={{ background: isHighlighted ? 'lightgray' : 'white' }}>
    <img width={30} height={30} src={item.url} ></img>  {" "+item.label}
    </div>
  }
  value={this.state.filtertext}
  onChange={(e) =>  this.setState({ filtertext: e.target.value })}
  onSelect={(val) => this.setState({ filtertext: val.charAt(0).toUpperCase()+val.slice(1) })}

/>
                    <PrimaryButton
                      text="filtrer"
                      style={{ float: "right" }}
                      onClick={() => this._onFilter(this.state.filtertext)}
                    />
                  </div>
                );
              }
            },
            {
              key: "Filter",
              name: "par",
              icon: "New",
              ariaLabel: "New. Use left and right arrow keys to navigate",
              onClick: () => {
                return;
              },
              subMenuProps: {
                items: [
                  {
                    key: "ID",
                    name: "ID",
                    icon: "Emoji",
                    onClick: () => {
                      this.setState({
                        filter: "ID"
                      });
                      return;
                    }
                  },
                  {
                    key: "Formations",
                    name: "Formations",
                    icon: "Education",
                    onClick: () => {
                      this.setState({
                        filter: "Formations"
                      });
                      return;
                    }
                  },
                  {
                    key: "Skills",
                    name: "Skills",
                    icon: "Crown",
                    onClick: () => {
                      this.setState({
                        filter: "Skills"
                      });
                      return;
                    }
                  },
                  {
                    key: "Expériences",
                    name: "Expériences",
                    icon: "Trophy",
                    onClick: () => {
                      this.setState({
                        filter: "Expériences"
                      });
                      return;
                    }
                  },
                  {
                    key: "Loisirs",
                    name: "Loisirs",
                    icon: "Soccer",
                    onClick: () => {
                      this.setState({
                        filter: "Loisirs"
                      });
                      return;
                    }
                  }
                ]
              }
            }
          ]}
          onRenderOverflowButton={this._onRenderOverflowButton}
          onRenderItem={this._onRenderItem}
        />
       </section> 
          <ToggleButtonGroup
            style={{ paddingRight: 290, paddingLeft: 290 }}
            value={_ComponenetContext.state.alignment}
            exclusive
            onChange={() =>
              this.getCurrentUser(
                _ComponenetContext,
                _ComponenetContext.state.token,
                true
              )
            }
          >
            <ToggleButton value="left">
              <StarRateIcon
                style={{ float: "left", width: "auto" }}
                fontSize={"large"}
              />
            </ToggleButton>
            <ToggleButton
              style={{ float: "right", width: "auto" }}
              value="right"
            >
              <SupervisorAccountIcon fontSize={"large"} />
            </ToggleButton>
          </ToggleButtonGroup>
        </div>
        <h1 style={{ textAlign: "center", color: "blue" }}>Liste des cv</h1>
        
        <Slider {...settings} id="slider">
          {this.Slide()}
        </Slider>
        

        <br />
        <br />

        <Panel
          isOpen={this.state.showPanel}
          onDismiss={this._hidePanel}
          type={PanelType.extraLarge}
          headerText="Contact"
          closeButtonAriaLabel="Close"
        >
          <Fabric className={SectionClassName}>
            <ComboBox
              required
              onChanged={this._onChange.bind(this)}
              label="Type de besoin"
              key={"" + state.autoComplete + state.allowFreeform}
              allowFreeform={state.allowFreeform}
              autoComplete={state.autoComplete ? "on" : "off"}
              options={ComboBoxOptions}
            />{" "}
            <span
              onClick={() => {
                this.setState({ allowFreeform: !this.state.allowFreeform });
              }}
            >
              <Toggle label="FreeForm" checked={state.allowFreeform} />
            </span>
            <span
              onClick={() => {
                this.setState({ autoComplete: !this.state.autoComplete });
              }}
            >
              <Toggle label="Auto-complete" checked={state.autoComplete} />
            </span>
          </Fabric>
          <TextField
            required
            label="Besoin"
            multiline
            rows={8}
            value = {this.state.besoin}
              onBlur={(option) => this.setState({ besoin: option.target.value })}
            iconProps={{ iconName: "glasses" }}
          />
          <section style={{float : 'left'}} >
<TimeKeeper
                  onChange={this.handleTimeChange}
                   time={this.state.clockTime}
                    config={{
                        useCoarseMinutes: true,
                        CLOCK_WRAPPER_MERIDIEM_BACKGROUND : '#F2F2F2',
                        CLOCK_WRAPPER_MERIDIEM_COLOR : '#F2F2F2'
                    }}
                />
</section>
<section style={{float : 'left'}} >
          <DayPicker
            selectedDays={this.state.selectedDays}
            onDayClick={this.handleDayClick.bind(this)}
            fromMonth={new Date()}
            disabledDays={disabled}
            toMonth={ this.getDateAfter2months() } 
          />
    </section>   
    <section style={{float : 'right' , color : 'black' }}   >
         {this.state.selectedDays.length>0 ?<div style={{ padding: 10 , border: "3px solid black", marginRight: 175, marginTop : 10 }} ><h4 style={{ color : '#E97ED4'  }} > Les dates selectionée(s) : </h4> {this.SelectedDates()} </div> :<div></div>}
    </section>  
         <br/>
         <br/>
         <br/>
         <br/>
         <br/>
         <br/>
         <br/>
         <br/>
    
        
          <MessageBar
            onDismiss={this._hidePanel}
            dismissButtonAriaLabel="Close"
            messageBarType={MessageBarType.warning}
            ariaLabel="Aria help text here"
            actions={
              <div>
                <MessageBarButton onClick={() => this.Entretien(this,this.state.token)}>
                  Demander un entretien
                </MessageBarButton>
              
              </div>
            }
          >
            J'ai lu et j'accepte le règlement du societé .
            <br />
            <Link href="https://www.journaldunet.fr/management/guide-du-management/1200721-l-entretien-professionnel-une-nouvelle-obligation/">le règlement du societé </Link>
          </MessageBar>
        </Panel>
        <Panel
          isOpen={this.state.showPanel2}
          onDismiss={this._hidePanel2}
          type={PanelType.extraLarge}
          headerText="Sign UP"
          closeButtonAriaLabel="Close"
        >
          <form
            onSubmit={() =>
              this.addUser(
                this,
                this.state.token,
                this.state.identifiant,
                this.state.company,
                this.state.password,
                this.state.adresse,
                this.state.telephone,
                this.state.nom,
                this.state.prenom
              )
            }
          >
            <TextField
              label="Email"
              type="email"
              iconProps={{ iconName: "mail" }}
              value={this.state.identifiant}
              onBlur={value =>
                this.setState({ identifiant: value.target.value })
              }
              validateOnLoad={false}
              onGetErrorMessage={this._getErrorMessage2}
              validateOnFocusIn
              validateOnFocusOut
              onRenderLabel={this._onRenderLabel.bind(this)}
              required
            />
            <TextField
              label="Nom"
              type="text"
              iconProps={{ iconName: "contactinfo" }}
              value={this.state.nom}
              onBlur={value => this.setState({ nom: value.target.value })}
              required
            />
            <TextField
              label="Prenom"
              type="text"
              iconProps={{ iconName: "contactinfo" }}
              value={this.state.prenom}
              onBlur={value => this.setState({ prenom: value.target.value })}
              required
            />
            <TextField
              label="adresse"
              type="text"
              iconProps={{ iconName: "backlogboard" }}
              value={this.state.adresse}
              onBlur={value => this.setState({ adresse: value.target.value })}
              required
            />
            <TextField
              label="Nom du societé"
              type="text"
              iconProps={{ iconName: "work" }}
              value={this.state.company}
              onBlur={value => this.setState({ company: value.target.value })}
              required
            />
            <TextField
              label="Mot de passe"
              type="password"
              iconProps={{ iconName: "shop" }}
              value={this.state.password}
              onBlur={value => this.setState({ password: value.target.value })}
              required
            />
            <TextField
              label="Confirmation mot de passe"
              type="text"
              iconProps={{ iconName: "shop" }}
              value={this.state.password2}
              onBlur={value => this.setState({ password2: value.target.value })}
              validateOnLoad={false}
              onGetErrorMessage={this._getErrorMessage}
              validateOnFocusIn
              validateOnFocusOut
              required
            />
            <br />
            <label style={{ fontStyle: "bold" }}>Numéro de telephone :</label>
            <ReactPhoneInput
              required
              defaultCountry={"tn"}
              value={this.state.telephone}
              onBlur={value => this.setState({ telephone: value.target.value })}
            />
            <br />
            <br /> <br />
            <br />
            <PrimaryButton
              type="submit"
              text="Sign UP"
              style={{ float: "right" }}
            />
          </form>
        </Panel>



        <Panel
          isOpen={this.state.showPanel3}
          onDismiss={this._hidePanel3}
          type={PanelType.extraLarge}
          headerText="Modifier le profil"
          closeButtonAriaLabel="Close"
        >
        <form onSubmit={()=>this.ChangeProfil(this.state.token)} >
            <Toggle 
            label= {'Changer le mot de passe '}
            style={{float : 'right'}}
            onClick={()=>this.setState({passwordChange: !this.state.passwordChange})}
            />
            <TextField
            disabled
              label="Email"
              type="email"
              iconProps={{ iconName: "mail" }}
              value={this.state.loggedinUser}
            />
            <TextField
              label="Nom"
              type="text"
              iconProps={{ iconName: "contactinfo" }}
              value={this.state.nom}
              onBlur={value => this.setState({ nom: value.target.value })}
              required
            />
            <TextField
              label="Prenom"
              type="text"
              iconProps={{ iconName: "contactinfo" }}
              value={this.state.prenom}
              onBlur={value => this.setState({ prenom: value.target.value })}
              required
            />
            <TextField
              label="adresse"
              type="text"
              iconProps={{ iconName: "backlogboard" }}
              value={this.state.adresse}
              onBlur={value => this.setState({ adresse: value.target.value })}
              required
            />
            <TextField
              label="Nom du societé"
              type="text"
              iconProps={{ iconName: "work" }}
              value={this.state.company}
              onBlur={value => this.setState({ company: value.target.value })}
              required
            />
            { this.state.passwordChange ?
              <div>
            <TextField
              label="Mot de passe"
              type="password"
              iconProps={{ iconName: "shop" }}
              value={this.state.password}
              onBlur={value => this.setState({ password: value.target.value })}
              required
            />
            <TextField
              label="Confirmation mot de passe"
              type="text"
              iconProps={{ iconName: "shop" }}
              value={this.state.password2}
              onBlur={value => this.setState({ password2: value.target.value })}
              validateOnLoad={false}
              onGetErrorMessage={this._getErrorMessage}
              validateOnFocusIn
              validateOnFocusOut
              required
            /> </div>: <div></div>
            
            }
            <br />
            <label style={{ fontStyle: "bold" }}>Numéro de telephone :</label>
            <ReactPhoneInput
              required
              defaultCountry={"tn"}
              value={this.state.telephone}
              onBlur={value => this.setState({ telephone: value.target.value })}
            />
            <br />
            <br /> <br />
            <br />
            <PrimaryButton
            type='submit'
            
              text="Modifier le profil"
              style={{ float: "right" }}
            />
         </form>
        </Panel>

        <Dialog
          hidden={this.state.hideDialog}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: "Connexion",
            subText:
              "Veuillez saisir votre identifiant et mot de passe pour se connecter"
          }}
          modalProps={{
            isBlocking: false,
            
          }}
        >
          <TextField
            label="Identifiant"
            type="text"
            iconProps={{ iconName: "contactinfo" }}
            value={this.state.login}
            onBlur={value => this.setState({ login: value.target.value })}
            onRenderLabel={this._onRenderLabel.bind(this)}
            required
          />
          <TextField
            label="Mot de passe"
            type="password"
            iconProps={{ iconName: "backlogboard" }}
            value={this.state.pw}
            onBlur={value => this.setState({ pw: value.target.value })}
            required
          />
          <DialogFooter>
            <PrimaryButton onClick={() => this.signin(this)} text="se connecter" />
            <DefaultButton onClick={this._closeDialog.bind(this)} text="annuler" />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  Entretien(componentContext,token): void {
    if(this.state.selectedDays.length >0){
      var EntretienFields = {
        fields: {
          Title: this.state.loggedinUser,
          typeBesoin: this.state.typebesoin,
          besoin: this.state.besoin,
          
          EmployeeMail : this.state.userPrincipalNameClicked
        }
      };
      $.ajax({
        async: true,
        crossDomain: true,
        url:
          "https://graph.microsoft.com/beta/sites/root/lists/25a45bd8-1c71-45a8-8980-1066cefebc74/items/",
        method: "POST",
        headers: {
          "content-type": "application/json"
        },
        data: JSON.stringify(EntretienFields),
        beforeSend: function(xhr) {
          xhr.setRequestHeader("Authorization", "Bearer " + token);
        },
        success: function(res) {

componentContext.state.selectedDays.forEach(element => {
        let x= new Date(element);
        let HeureText = ""+(x.getHours()-1).toString()+":"+x.getMinutes()
        let DateText = ""+x.getFullYear()+"-"+(x.getMonth()+1).toString()+"-"+x.getDay()
        
            var DateEntretien= {
              fields: {
              Title: res.fields.id,
              NumEntretienLookupId: res.fields.id ,
              Date : DateText,
              HeureEntretien: HeureText
              }
            };
            $.ajax({
              async: true,
              crossDomain: true,
              url:
                "https://graph.microsoft.com/beta/sites/root/lists/bd317c32-a24d-4d3e-90ae-46434ae3c668/items/",
              method: "POST",
              headers: {
                "content-type": "application/json"
              },
              data: JSON.stringify(DateEntretien),
              beforeSend: function(xhr2) {
                xhr2.setRequestHeader("Authorization", "Bearer " + token);
              },
              success: function() {
    
              }
            });
          });

          
        }
      });
   
    this._hidePanel()
    alert("Demande envoyée.")
    }
    else
    {
      alert("Veuillez selectionner au moins une date valide.")
    }
  }

  private _onChange = (option?: IComboBoxOption, index?: number, value?: string): void => {
   
    if (option) {
   
      this.setState({
        typebesoin: option.text
      });
    } else if (value !== undefined) {
      const newOption: IComboBoxOption = { key: value, text: value };
    
      this.setState({
       
        typebesoin: newOption.text
      });
    }
  };

  SelectedDates(): any {
    let x =[]
    this.state.selectedDays.forEach(element => {
      x.push([<p>{" DATE : "+element.toLocaleDateString()+" Heure : "+element.toLocaleTimeString()}</p>])
    });
    return x
  }

private getSuggeestions() : any
{
      switch (this.state.filter)
      {
        case "ID": {
          return[];
          break;
        }
        case "Formations": {
          return [];
          break;
        }
  
        case "Skills": {
        let x=[{ label: 'Java',url:'https://upload.wikimedia.org/wikipedia/fr/2/2e/Java_Logo.svg' },
        { label: 'React',url:'https://upload.wikimedia.org/wikipedia/commons/a/a7/React-icon.svg'},
        { label: 'HTML/CSS' ,url :'https://upload.wikimedia.org/wikipedia/commons/d/d5/CSS3_logo_and_wordmark.svg'},
        { label: 'TypeSript' , url:'https://www.vectorlogo.zone/logos/typescriptlang/typescriptlang-icon.svg' }]

     return x
    break;
        }
  
        case "Expériences": {
          let x=[{ label: 'Airbnb',url:'https://img.icons8.com/color/48/000000/airbnb.png' },
          { label: 'Buffer',url:'https://img.icons8.com/doodle/48/000000/buffer.png'},
          { label: 'Financer' ,url :'https://img.icons8.com/dusk/64/000000/e-commerce.png'},
          { label: 'Port-Finder' ,url :'https://img.icons8.com/dusk/64/000000/password.png'}
        ]
        return x
          break;
        }
  
        case "Loisirs": {
          let x=[{ label: 'sport',url:'https://img.icons8.com/cotton/64/000000/trainers.png' },
          { label: 'musique',url:'https://img.icons8.com/nolan/64/000000/musical-notes.png'},
          { label: 'films' ,url :'https://img.icons8.com/dusk/64/000000/film-reel.png'},]
          
  
       return x
          break;
        }
  
        default: return null
          break;
        
      }

  
}
  private getDateAfter2months():Date
  {
   let x = new Date()
 x.setMonth(x.getMonth()+2);
return x
  }
 private ChangeProfil(token): void {
  
  var _secretKey = "elomri";
  var simpleCrypto = new SimpleCrypto(_secretKey);
  var cryptedpassword = simpleCrypto.encrypt(this.state.password);
  var x = {
    
      
      company: this.state.company,
      password: cryptedpassword,
      adresse: this.state.adresse,
      telephone: this.state.telephone,
      nom: this.state.nom,
      email: this.state.prenom,
      
    
  };
  $.ajax({
    async: true,
    crossDomain: true,
    url:
      "https://graph.microsoft.com/beta/sites/root/lists/cec630c7-c1f1-4025-a8b2-d77167035e5d/items/"+this.state.id+"/fields",
    method: "PATCH",
    headers: {
      "content-type": "application/json"
    },
    data: JSON.stringify(x),
    beforeSend: function(xhr) {
      xhr.setRequestHeader("Authorization", "Bearer " + token);
    },
    success: function() {}
  });
  
  }

  handleTimeChange(newTime){

    let x ={
      hour : newTime.hour,
      minute: newTime.minute
    }
    this.setState({ clockTime: x})
}
  
}
