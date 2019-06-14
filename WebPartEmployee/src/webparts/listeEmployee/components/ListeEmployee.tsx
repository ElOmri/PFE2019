import * as React from "react";
import { sp } from "@pnp/sp";
import { DateRange } from 'react-date-range';
import { CurrentUser } from '@pnp/sp/src/siteusers';
import { IListeEmployeeProps } from "./IListeEmployeeProps";
import { format, addDays } from 'date-fns';
import { DateRangePicker } from 'react-date-range';
import 'react-date-range/dist/styles.css'; // main style file
import 'react-date-range/dist/theme/default.css'; // theme css file
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { TeachingBubble } from "office-ui-fabric-react/lib/TeachingBubble";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { DefaultButton, PrimaryButton } from "office-ui-fabric-react";
import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import BootstrapTable from "react-bootstrap-table-next";
import "react-bootstrap-table-next/dist/react-bootstrap-table2.min.css";
import ToolkitProvider, { Search } from "react-bootstrap-table2-toolkit";
import paginationFactory from "react-bootstrap-table2-paginator";
import { CSVExport } from "react-bootstrap-table2-toolkit";
import "bootstrap/dist/css/bootstrap.min.css";
const { ExportCSVButton } = CSVExport;
const { SearchBar } = Search;
export interface IState {
  idselected:string;
  infoClient:any;
  CurrentUser:string;
  Entretiens: any;
  DateEntretien: any;
  isTeachingBubbleVisible?: boolean;
  verif: boolean;
  showPanel: boolean;
  Client: string;
  idEntretien: string;
  toggle:boolean;
  selectedDate:string;
  emailValue:string;
  dateRange:any;
}

const sizePerPageRenderer = ({
  options,
  currSizePerPage,
  onSizePerPageChange
}) => (
  <div className="btn-group" role="group">
    {options.map(option => {
      const isSelect = currSizePerPage === `${option.page}`;
      return (
        <button
          key={option.text}
          type="button"
          onClick={() => onSizePerPageChange(option.page)}
          className={`btn ${isSelect ? "btn-secondary" : "btn-warning"}`}
        >
          {option.text}
        </button>
      );
    })}
  </div>
);

const options = {
  sizePerPageRenderer
};

const columns = [
  {
    dataField: "Id",
    text: "Id Entretien",
    sort: true
  },
  {
    dataField: "Client",
    text: "Email de client",
    sort: true
  },
  {
    dataField: "besoinType",
    text: "Type de besoin",
    sort: true
  },
  {
    dataField: "besoin",
    text: "Besoin",
    sort: true
  }
];

export default class ListeEmployee extends React.Component<
  IListeEmployeeProps,
  IState
> {
  constructor(props: IListeEmployeeProps) {
    super(props);
    this.state = {
      dateRange: {
        selection: {
          startDate: new Date(),
          endDate: addDays(new Date(), 30),
          key: 'selection',
        },
      },
      emailValue:"Bonjour ,\nEntretien est confirmé le DATE_DEBUT jusqu'a DATE_FIN,\nCordialement. ",
      idselected:"",
      selectedDate:"",
      infoClient :
      {
        mail:"User@gmail.com",
        company: "proged",
        adresse : "riadh",
        telephone :"+21650699436",
        prenom : "anas",
        nom : "omri"
      },
      CurrentUser:null,
      toggle:false,
      DateEntretien: { key: "DEFAULT", text: "DEFAULT_DATE" },
      Client: "",
      idEntretien: "1",
      showPanel: false,
      verif: false,
      Entretiens: null,
      isTeachingBubbleVisible: false
    };
  }

  private _showPanel(id) {
   setTimeout(() => {
    this.setState({ showPanel: true });
   }, 750); 
    this.ListeDate(id);
  }

  private _hidePanel = (): void => {
    this.setState({ showPanel: false });
  };

  private _onDismiss(): void {
    this.setState({
      isTeachingBubbleVisible: false
    });
  }

  public componentDidMount(): void {
    this.getCurrentUser();
    setTimeout(() => {
      this.ListeEntretien();
    }, 1000);
   
  }
  private getCurrentUser(): void {    
    sp.web.currentUser.get().then((r: CurrentUser) => {
      let currentUser =r['Email'];
      alert("Bonjour "+currentUser)
      this.setState({ CurrentUser: currentUser });
    });
  }

  public rowEvents = {
    onMouseEnter: (e, row, rowIndex) => {
      if(!this.state.verif){
      this.setState({verif:true})
      this.InfoUser(row.Client)
      setTimeout(() => {
        
        this.setState({verif:false,isTeachingBubbleVisible: true});
      }, 750);
    
      }
    },
    onMouseLeave: (e, row, rowIndex) => {

      
      this.setState({isTeachingBubbleVisible: false});
    
      
    },
    onClick: (e, row) => {
      this._showPanel(row.Id);
      this.setState({
        Client: row.Client,
        
        
        idEntretien: row.Id
      });
    }
  };

  private ListeEntretien(): void {
    let texto = "sites/root/lists/25a45bd8-1c71-45a8-8980-1066cefebc74/items/";
    this.props.context.msGraphClientFactory.getClient().then(client => {
      client
        .api(texto)
        .version("beta")
        .expand("fields")
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .filter("fields/EmployeeMail eq '"+this.state.CurrentUser+"' and  fields/Effectue eq false ")
        .get((err, res) => {
          if (err) {
            console.error(err);
            return;
          }

          let x = res.value.map(item => ({
            Id: item.fields.id,
            Client: item.fields.Title,
            besoinType: item.fields.typeBesoin,
            besoin: item.fields.besoin
          }));
          this.setState({
            Entretiens: x
          });
          console.log(res);
        });
    });
  }

  private ListeDate(idEntretien): void {
    let texto = "sites/root/lists/bd317c32-a24d-4d3e-90ae-46434ae3c668/items/";
    this.props.context.msGraphClientFactory.getClient().then(client => {
      client
        .api(texto)
        .version("beta")
        .expand("fields")
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .filter("fields/NumEntretienLookupId eq '" + idEntretien + "'")
        .get((err, res) => {
          if (err) {
            console.error(err);
            return;
          }

          let x = res.value.map(item => ({
            key: item.fields.Date,
            text:
             
              item.fields.Date +
              " " +
              item.fields.HeureEntretien+":0"
          }));
          this.setState({
            DateEntretien: x
          });
          console.log(res);
        });
    });
  }


  private InfoUser(Client): void {
    let texto = "sites/root/lists/cec630c7-c1f1-4025-a8b2-d77167035e5d/items/";
    this.props.context.msGraphClientFactory.getClient().then(client => {
      client
        .api(texto)
        .version("beta")
        .expand("fields")
        .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
        .filter("fields/Title eq '"+ Client +"'")
        .get((err, res) => {
          if (err) {
            console.error(err);
            return;
          }

          let x = {
            mail: res.value[0].fields.Title,
            company: res.value[0].fields.company,
            nom: res.value[0].fields.nom,
            prenom: res.value[0].fields.email,
            adresse: res.value[0].fields.adresse,
            telephone: res.value[0].fields.telephone
          }
          this.setState({
            infoClient: x
          });
          console.log(res);
        });
    });
  }


  public render(): React.ReactElement<IListeEmployeeProps> {
    if (this.state.Entretiens == null) {
      return <div>Loading ..</div>;
    }
    return (
      <div>
        <div>Liste des rendez-vous :</div>
        <ToolkitProvider
          keyField="id"
          data={this.state.Entretiens}
          columns={columns}
          search
        >
          {props => (
            <div>
              <SearchBar {...props.searchProps} />
              <hr />
              <div id="table" ></div>
              <BootstrapTable
                className={".table-striped"}
                {...props.baseProps}
                rowEvents={this.rowEvents}
                bootstrap4
                pagination={paginationFactory(options)}
              />
              
              <ExportCSVButton {...props.csvProps}>
                Exporter les entretiens en excel
              </ExportCSVButton>
            </div>
          )}
        </ToolkitProvider>

        {this.state.isTeachingBubbleVisible ? (
          <div>
            <TeachingBubble
           
              onDismiss={this._onDismiss.bind(this)}
              hasCloseIcon={true}
              headline="informations client"
            >
             
              <ul> 
<li>Email : {this.state.infoClient.mail}</li>
<li>Societé : {this.state.infoClient.company}</li>
<li>Adresse : {this.state.infoClient.adresse}</li>
<li>Nom : {this.state.infoClient.nom}</li>
<li>Prenom : {this.state.infoClient.prenom}</li>
<li>Telephone :{this.state.infoClient.telephone}</li>
              </ul>
            </TeachingBubble>
          </div>
        ) : null}

        <div>
          <Panel
            isOpen={this.state.showPanel}
            onDismiss={this._hidePanel}
            type={PanelType.medium}
            headerText="Approuver un entretien"
            closeButtonAriaLabel="Close"
          >
            <form onSubmit={() => this.Effectuer()} >
             <Toggle
          defaultChecked={true}
          label="
          Personnaliser l'email."
          onText="On"
          offText="Off"
         
          onChange={()=>this.setState({toggle:!this.state.toggle})}
        />
            <Dropdown
            required
              label="Date Entretien"
              options={this.state.DateEntretien}
              onChanged={ this._onDropdownChanged.bind(this) } 
            />
            {(this.state.toggle)?
            <TextField required label="Email :" multiline rows={5} value={this.state.emailValue} onChange={(x,value)=>this.setState({ emailValue:value})} />:<div></div>
            }
            <h2>délai de mission :</h2>
            <br/>
 <DateRange
           onChange={this.handleSelect.bind(this)}
            moveRangeOnFirstSelection={false}
            ranges={[this.state.dateRange.selection]}
            minDate={new Date()}
            className={'PreviewArea'}
          />
<br/>
<section style={{ float : 'right', margin : 50}}>
            <PrimaryButton
              text="Effectuer entretien"
              type="submit"
              disabled={this.state.selectedDate==""}
            />
            <DefaultButton text="Annuler" onClick={() => this._hidePanel()} />
            </section>
            </form>

          </Panel>
        </div>
      </div>
    );
  }


  
  Effectuer(): void {
      if(this.state.selectedDate!="")
      {
            let datex=new Date(this.state.selectedDate);
            let dateend= new Date(datex.setHours(datex.getHours()+1))
          
            let body={
              "subject": "Entretien",
              "start": {
                "dateTime": datex.toUTCString(),
                "timeZone": "UTC"
              },
              "end": {
                "dateTime": dateend.toUTCString(),
                "timeZone": "UTC"
              }
            }
            let event={
              "subject": "Mission",
              "start": {
                "dateTime": this.state.dateRange.selection.startDate.toUTCString(),
                "timeZone": "UTC"
              },
              "end": {
                "dateTime": this.state.dateRange.selection.endDate.toUTCString(),
                "timeZone": "UTC"
              }
            }

            let message={
              
                "message": {
                  "subject": "Entretien",
                  "body": {
                    "contentType": "Text",
                    "content": "Bonjour ,\nEntretien est confirmé le "+datex.toString() +" jusqu'a "+dateend.toString()+". \n Cordialement."
                  },
                  "toRecipients": [
                    {
                      "emailAddress": {
                        "address": this.state.Client
                      }
                    }
                  ]
                }
              }
              let message2={
              
                "message": {
                  "subject": "Entretien",
                  "body": {
                    "contentType": "Text",
                    "content": this.state.emailValue
                  },
                  "toRecipients": [
                    {
                      "emailAddress": {
                        "address": this.state.Client
                      }
                    }
                  ]
                }
              }



            let textevent = "me/events";
            let textMessage="me/sendMail"

    this.props.context.msGraphClientFactory.getClient().then(client => {
      client
        .api(textevent)
        .version("beta")
        .header("Content-type", "application/json")
        
        .post(body, (err, res) => {
          if (err) {
            console.error(err);
            return;
          }
         
        });
    });

    this.props.context.msGraphClientFactory.getClient().then(client => {
      client
        .api(textevent)
        .version("beta")
        .header("Content-type", "application/json")
        
        .post(event, (err, res) => {
          if (err) {
            console.error(err);
            return;
          }
         
        });
    });
if(!this.state.toggle){
    this.props.context.msGraphClientFactory.getClient().then(client => {
      client
        .api(textMessage)
        .version("beta")
        .header("Content-type", "application/json")
        
        .post(message, (err, res) => {
          if (err) {
            console.error(err);
            return;
          }
     
        });
    });
  }else
  {
    this.props.context.msGraphClientFactory.getClient().then(client => {
      client
        .api(textMessage)
        .version("beta")
        .header("Content-type", "application/json")
        
        .post(message2, (err, res) => {
          if (err) {
            console.error(err);
            return;
          }
       
        });
    });
  }


    this.props.context.msGraphClientFactory.getClient().then(client => {
      client
        .api(textMessage)
        .version("beta")
        .header("Content-type", "application/json")
        
        .post(message, (err, res) => {
          if (err) {
            console.error(err);
            return;
          }
         
        });
    });
    let x={
      "Effectue" : true
    }
    let textpatch = "sites/root/lists/25a45bd8-1c71-45a8-8980-1066cefebc74/items/"+this.state.idEntretien+"/fields";
    this.props.context.msGraphClientFactory.getClient().then(client => {
      client
        .api(textpatch)
        .version("beta")
        
        
        .patch(x,(err, res) => {
          if (err) {
            console.error(err);
            return;
          }

         
         
        });
    });
            alert("Entretien Effectué.")
      }else
      {
        alert("selectionner une date")
      }

  }

  _onDropdownChanged(event) { 
    var newValue = event.text; 
    this.setState( { selectedDate: newValue} ); 
    
}

private handleSelect(range){
  console.log(range)
 this.setState({
   dateRange : range
 })
}
}
