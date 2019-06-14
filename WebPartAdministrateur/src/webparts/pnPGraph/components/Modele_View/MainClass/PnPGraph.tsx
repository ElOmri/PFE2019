import { Client } from "@microsoft/microsoft-graph-client";
import { User as IUser } from "@microsoft/microsoft-graph-types-beta";
import { MSGraphClient } from "@microsoft/sp-http";
import { escape } from "@microsoft/sp-lodash-subset";
import { graph } from "@pnp/graph";
import { CurrentUser } from "@pnp/sp/src/siteusers";
import { CommandBarButton } from "office-ui-fabric-react/lib/Button";
import {
  DetailsList,
  DetailsListLayoutMode,
  IDetailsList,
  IColumn,
  Selection
} from "office-ui-fabric-react/lib/DetailsList";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import { Image, ImageFit } from "office-ui-fabric-react/lib/Image";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import {
  IOverflowSetItemProps,
  OverflowSet
} from "office-ui-fabric-react/lib/OverflowSet";
import { Panel, PanelType } from "office-ui-fabric-react/lib/Panel";
import { ITag, TagPicker } from "office-ui-fabric-react/lib/Pickers";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { createRef } from "office-ui-fabric-react/lib/Utilities";
import * as React from "react";
import LoadingScreen from "react-loading-screen";
import pnp from "sp-pnp-js";
import { IDetailsListBasicExampleState } from "../../Models/IDetailsListBasicExampleState";
import { IPnPGraphProps } from "../../Models/IPnPGraphProps";
import { IState } from "../../Models/IState";
import { IUserItem } from "../../Models/IUserItem";
import { Template } from "../Template/template";
import { _columns } from "../../DataProvider/DetailListColumns";
import { OverflowSetStyle } from "../../DataProvider/OverflowStyle";

export default class PnPGraph extends React.Component<
  IPnPGraphProps,
  IState,
  IDetailsListBasicExampleState
> {
  private _selection: Selection;
  private resume: any;
  private _detailsList = createRef<IDetailsList>();
  private _testTags: ITag[];
  private Client: Client;
  image;
  pdfExportComponent;

  constructor(props: IPnPGraphProps, state: IState) {
    super(props);
    this._selection = new Selection({
      onSelectionChanged: () =>
        this.setState({ selectionDetails: this._getSelectionDetails() })
    });
    this.state = {
      events:null,
      projects: null,
      templateFile: "",
      selectionDetails: this._getSelectionDetails(),
      libraries: null,
      showPanel: false,
      users: null,
      filter: "name",
      hide: false,
      displayName: null,
      mail: null,
      userPrincipalName: null,
      businessPhones: null,
      city: null,
      country: null,
      officeLocation: null,
      streetAddress: null,
      skills: null,
      schools: null,
      userclicked: "",
      urlsite: "cvgenerator.sharepoint.com",
      aboutMe: null,
      interests: null,
      pastProjects: null,
      companyName: null,
      department: null,
      jobTitle: null,
      responsibilities: null,
      favoriImage: true,
      userfav: []
    };
  }

  private getListUserFav(): Boolean {
    pnp.sp.web.lists
      .getByTitle("FavoriUsers")
      .items.select("Title")
      .get()
      .then(r => {
        for (var i = 0; i < r.length; i++) {
          //console.log(r[i].Title)
          this.state.userfav.push(r[i].Title);
        }
      });
    console.log(this.state.userfav);
    if (this.state.userfav) return true;
    else return false;
  }
  
  private onclickFavoriImage() {
    if (this.state.favoriImage) {
      if (this.getListUserFav()) {
        let x;
        x = [];
        graph.users.get<IUser[]>().then(users => {
          users.forEach(element => {
            if (this.state.userfav.includes(element.userPrincipalName)) {
              x.push(element);
            }
          });
        });

        this.setState({
          users: x
        });
      }
      this.setState({ favoriImage: false, filter: "favori" });
    } else {
      this.setState({ favoriImage: true, filter: "name", userfav: [] });

      graph.users.get<IUser[]>().then(users => {
        this.setState({
          users
        });
      });
    }
  }
  private verifFiles() {
    pnp.sp.web.lists
      .getById(this.props.Lists)
      .items.filter("FSObjType eq 0")
      .select("FileLeafRef,FileRef")
      .get()
      .then(files => {
        files.forEach(f => {
          if (this.props.TemplateFile == f.FileLeafRef) {
            alert("upload");
          }
        });
      });
  }
  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();
    var i = "";
    for (var _i = 0; _i < selectionCount; _i++) {
      i += "" + (this._selection.getSelection()[_i] as IUserItem).displayName;
    }
    return i + " items selected";
  }

  private getselection(): ITag[] {
    let x: ITag[] = (this._selection.getSelection() as IUserItem[]).map(
      item => ({ key: item.displayName, name: item.displayName })
    );
    return x;
  }

  public componentDidMount(): void {
    graph.users.get<IUser[]>().then(users => {
      this.setState({
        users
      });
      this._testTags = this.state.users.map(item => ({
        key: item.displayName,
        name: item.displayName
      }));
    });
    this.verifFiles();
  }

  private _cvGenerator = function() {
    this._searchWithGraph(this.state.userclicked);
    return (
      <div>
        <Template
        projects={this.state.projects}
          templatemodele={this.state.templateFile}
          responsibilities={this.state.responsibilities}
          displayName={this.state.displayName}
          mail={this.state.mail}
          userPrincipalName={this.state.userPrincipalName}
          businessPhones={this.state.businessPhones}
          city={this.state.city}
          country={this.state.country}
          officeLocation={this.state.officeLocation}
          streetAddress={this.state.streetAddress}
          skills={this.state.skills}
          schools={this.state.schools}
          aboutMe={this.state.aboutMe}
          interests={this.state.interests}
          pastProjects={this.state.pastProjects}
          companyName={this.state.companyName}
          department={this.state.department}
          jobTitle={this.state.jobTitle}
        />
      </div>
    );
  };

  private _onFilter(text: string): void {
    console.log(this.state.users);

    if (!text) {
      graph.users.get<IUser[]>().then(users => {
        this.setState({
          users
        });
      });
    }
    switch (this.state.filter) {
      case "name": {
        if (text) {
          this.setState({
            users: this.state.users.filter(i => {
              if (i.displayName != null)
                return i.displayName.toLowerCase().indexOf(text) > -1;
            })
          });
          break;
        } else {
          graph.users.get<IUser[]>().then(users => {
            this.setState({
              users
            });
          });
          break;
        }
      }
      case "Departement": {
        if (text) {
          this.setState({
            users: this.state.users.filter(i => {
              if (i.officeLocation != null)
                return i.officeLocation.toLowerCase().indexOf(text) > -1;
            })
          });
          break;
        } else {
          graph.users.get<IUser[]>().then(users => {
            this.setState({
              users
            });
          });
          break;
        }
      }
      case "job": {
        if (text) {
          this.setState({
            users: this.state.users.filter(i => {
              if (i.jobTitle != null)
                return i.jobTitle.toLowerCase().indexOf(text) > -1;
            })
          });
          break;
        } else {
          graph.users.get<IUser[]>().then(users => {
            this.setState({
              users
            });
          });
          break;
        }
      }
      case "userPrincipalName": {
        if (text) {
          this.setState({
            users: this.state.users.filter(i => {
              if (i.userPrincipalName != null)
                return i.userPrincipalName.toLowerCase().indexOf(text) > -1;
            })
          });
          break;
        } else {
          graph.users.get<IUser[]>().then(users => {
            this.setState({
              users
            });
          });
          break;
        }
      }
      default:
        this.setState({
          users: text
            ? this.state.users.filter(
                i => i.displayName.toLowerCase().indexOf(text) > -1
              )
            : this.state.users
        });
        break;
    }
  }

  private ListeProjets(): void {
    let texto = "sites/root/lists/94fa4f56-2957-4d33-a71d-c8c8488469cf/items/";
    this.props.context.msGraphClientFactory.getClient().then(
      (client: MSGraphClient): void => {
        client
          .api(texto)

          .version("beta")
          
              .expand("fields")
          .get((err, res) => {
            if (err) {
              console.error(err);
              return;
            }
          this.setState({
            projects: res
          })
          console.log(this.state.projects)
          });
      }
    );
  }

  private _searchWithGraph(text: string): void {
    if(text){
    let texto = "users/" + text;
    this.props.context.msGraphClientFactory.getClient().then(
      (client: MSGraphClient): void => {
        client
          .api(texto)

          .version("beta")
          .select(
            "displayName,mail,userPrincipalName,businessPhones,city,country,officeLocation,streetAddress,skills,schools, aboutMe,interests,pastProjects,companyName,department,jobTitle,responsibilities"
          )
          // .filter(`(displayName eq '${escape(text)}')`)
          .get((err, res) => {
            if (err) {
              console.error(err);
              return;
            }
            this.setState({
              displayName: res.displayName,
              mail: res.mail,
              userPrincipalName: res.userPrincipalName,
              businessPhones: res.businessPhones,
              city: res.city,
              country: res.country,
              officeLocation: res.officeLocation,
              streetAddress: res.streetAddress,
              skills: res.skills,
              schools: res.schools,
              aboutMe: res.aboutMe,
              interests: res.interests,
              pastProjects: res.pastProjects,
              companyName: res.companyName,
              department: res.department,
              jobTitle: res.jobTitle,
              userclicked: null,
              responsibilities: res.responsibilities
            });
          });
      }
    );
    this.ListeProjets();
    this.events(text);
    }
  }

  private events(text: string): void {
    if(text){
    let texto = "users/" + text+"/events";
    this.props.context.msGraphClientFactory.getClient().then(
      (client: MSGraphClient): void => {
        client
          .api(texto)

          .version("beta")
          .select(
            "subject,body,start,end"
          )
          // .filter(`(displayName eq '${escape(text)}')`)
          .get((err, res) => {
            if (err) {
              console.error(err);
              return;
            }
            this.setState({
              events:res
            });
            console.log(this.state.events)
          });
      }
    );
   
    }
  }

  private _showPanel = (item: any): void => {
    pnp.sp.web.currentUser.get().then((r: CurrentUser) => {
      console.log(r);
    });
    //------------Get fileName ------------------------------------------------------------------------
    pnp.sp.web.lists
      .getById(escape(this.props.Lists))
      .items.filter("FSObjType eq 0")
      .select("FileLeafRef,FileRef")
      .get()
      .then(files => {
        let verif = false;
        files.forEach(f => {
          if (escape(this.props.TemplateFile) == f.FileLeafRef) {
            pnp.sp.web
              .getFileByServerRelativeUrl(f.FileRef)
              .getText()
              .then((text: string) => {
                this.setState({
                  templateFile: text
                });
              });

            setTimeout(() => {
              /*alert("Template Uploaded")*/
            }, 700);
            verif = true;
          }
        });
        if (!verif) {
          /*alert("file not found")*/
        }
      });
    setTimeout(() => {
      this.setState({
        showPanel: true
      });
    }, 2000);

    this.setState({
      //  showPanel: true,
      userclicked: item.userPrincipalName
    });
  };

  private _hidePanel = (): void => {
    this.setState({ showPanel: false });
  };

  private _renderItemColumn(item: IUser, index: number, column: IColumn) {
    const fieldContent = item[column.key as keyof IUser] as string;

    switch (column.key) {
      case "photo":
        return (
          <div
            style={{
              width: 100,
              height: 100,
              borderBottomLeftRadius: 50,
              borderBottomRightRadius: 50,
              borderTopRightRadius: 50,
              borderTopLeftRadius: 50,
              overflow: "hidden"
            }}
          >
            {" "}
            <Image
              src={"/_layouts/15/userphoto.aspx?size=L&username=" + item.mail}
              imageFit={ImageFit.cover}
              maximizeFrame={true}
            />
          </div>
        );

      case "name":
        return <span style={{ marginTop: "30px" }}>{item.displayName}</span>;

      case "jobTitle":
        var randomColor = "#" + (((1 << 24) * Math.random()) | 0).toString(16);
        return (
          <span style={{ color: randomColor, fontFamily: "Comic Sans MS" }}>
            {item.jobTitle}
          </span>
        );

      case "officeLocation":
        return <span>{item.officeLocation}</span>;

      case "userPrincipalName":
        return <span>{item.userPrincipalName}</span>;

      default:
        return <span>HI</span>;
    }
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

  private _getTextFromItem(item: ITag): string {
    return item.name;
  }

  private toggleHidden() {
    this.setState({
      hide: !this.state.hide
    });
  }

  private _onDisabledButtonClick = (): void => {
    this.setState({
      isPickerDisabled: !this.state.isPickerDisabled
    });
  };

  private _onFilterChanged = (filterText: string, tagList: ITag[]): ITag[] => {
    return filterText
      ? this._testTags.filter(
          tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0
        )
      : [];
  };

  private _onFilterChangedNoFilter = (
    filterText: string,
    tagList: ITag[]
  ): ITag[] => {
    return filterText
      ? this._testTags.filter(
          tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0
        )
      : [];
  };

  private _listContainsDocument(tag: ITag, tagList?: ITag[]) {
    if (!tagList || !tagList.length || tagList.length === 0) {
      return false;
    }
    return tagList.filter(compareTag => compareTag.key === tag.key).length > 0;
  }

  public render(): JSX.Element {
    const { favoriImage } = this.state;
    if (!this.state.users) {
      return (
        <LoadingScreen
          loading={true}
          bgColor="#f1f1f1"
          spinnerColor="#9ee5f8"
          textColor="#676767"
          logoSrc="../images/loader.gif"
          text="Loading Interface CV Generator"
        >
          <div />
        </LoadingScreen>
      );
    }

    return (
      <div>
        <h2>Liste des employés :</h2>
        <br />
        <br />
        <Fabric>
          <div>
            {favoriImage ? (
              <img
                onClick={this.onclickFavoriImage.bind(this)}
                src="https://image.flaticon.com/icons/svg/149/149222.svg"
                width="25px"
                height="25px"
                style={{ float: "left" }}
              />
            ) : (
              <img
                onClick={this.onclickFavoriImage.bind(this)}
                src="https://image.flaticon.com/icons/svg/148/148841.svg"
                width="25px"
                height="25px"
                style={{ float: "left" }}
              />
            )}
            <OverflowSet
              items={[
                {
                  key: "search",
                  onRender: () => {
                    return (
                      <TextField
                        placeholder={"Recherche par " + this.state.filter}
                        style={{ width: "600px", fontFamily: "Comic Sans MS" }}
                        className={OverflowSetStyle}
                        onChanged={this._onFilter.bind(this)}
                        underlined={true}
                        iconProps={{ iconName: "Filter" }}
                      />
                    );
                  }
                },
                {
                  key: "Filter",
                  name: "Filtrer",
                  icon: "New",
                  ariaLabel: "New. Use left and right arrow keys to navigate",
                  onClick: () => {
                    return;
                  },
                  subMenuProps: {
                    items: [
                      {
                        key: "DisplayName",
                        name: "Nom",
                        icon: "Emoji",
                        onClick: () => {
                          this.setState({
                            filter: "name"
                          });
                          return;
                        }
                      },
                      {
                        key: "Departement",
                        name: "Departement",
                        icon: "EMI",
                        onClick: () => {
                          this.setState({
                            filter: "Departement"
                          });
                          return;
                        }
                      },
                      {
                        key: "jobTitle",
                        name: "Fonction",
                        icon: "TeleMarketer",
                        onClick: () => {
                          this.setState({
                            filter: "job"
                          });
                          return;
                        }
                      },
                      {
                        key: "userPrincipalName",
                        name: "e-mail",
                        icon: "mail",
                        onClick: () => {
                          this.setState({
                            filter: "userPrincipalName"
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
          </div>
          <MarqueeSelection selection={this._selection}>
            <DetailsList
              componentRef={this._detailsList}
              items={this.state.users}
              columns={_columns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionPreservedOnEmptyClick={true}
              ariaLabelForSelectionColumn="Toggle selection"
              ariaLabelForSelectAllCheckbox="Toggle selection for all items"
              onItemInvoked={this._showPanel}
              selection={this._selection}
              onRenderItemColumn={this._renderItemColumn}
            />

            <TagPicker
              onResolveSuggestions={this._onFilterChanged}
              selectedItems={this.getselection()}
              getTextFromItem={this._getTextFromItem}
              pickerSuggestionsProps={{
                suggestionsHeaderText: "Suggested Tags",
                noResultsFoundText: "No Name Tags Found"
              }}
              itemLimit={10}
              disabled={false}
            />
          </MarqueeSelection>
        </Fabric>
        <div />
        <Panel
          isOpen={this.state.showPanel}
          onDismiss={this._hidePanel}
          type={PanelType.extraLarge}
          headerText={this.state.displayName}
          closeButtonAriaLabel="Close"
        >
          {this._cvGenerator()}
        </Panel>
      </div>
    );
  }
}
