import * as React from "react";
import * as $ from "jquery";
import {
  IPersonaProps,
  Persona,
  PersonaSize
} from 'office-ui-fabric-react/lib/Persona';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';

interface GraphFabricProps { }
interface IGraphFabricState {
    myData: IDisplayProps
}

interface IDisplayProps {
    displayName: string;
    givenName: string;
    jobTitle: string;
    mobilePhone: string;
    imageUrl: string;
}

export class GraphFabric extends React.Component<GraphFabricProps, IGraphFabricState> {
    constructor(props: GraphFabricProps) {
        super(props);
        this.state = {
            myData: {
                displayName: "",
                givenName: "",
                jobTitle: "",
                mobilePhone: "",
                imageUrl: ""
            }
        };
    }
    public componentDidMount(): void {
        
        this.getDetails().then(r => {
            this.getMyPhotoDef().then(i => {
                this.setState({
                    myData: {
                        displayName: r.displayName,
                        givenName: r.givenName,
                        jobTitle: r.jobTitle,
                        mobilePhone: r.mobilePhone,
                        imageUrl: i
                    }
                })
            })
        });

    }

    public render(): React.ReactElement<GraphFabricProps> {
        
        let examplePersona = {
            imageUrl: this.state.myData.imageUrl,
            imageInitials: '',
            primaryText: this.state.myData.displayName,
            secondaryText: this.state.myData.givenName,
            tertiaryText: this.state.myData.jobTitle,
            optionalText: this.state.myData.mobilePhone
          };
        return (
            <div>
                <Persona
                { ...examplePersona }
                size={ PersonaSize.extraLarge }
                coinSize={ 100 }
                onRenderSecondaryText={ this._onRenderSecondaryText } />
            </div>
        );
    }

    @autobind
    private _onRenderSecondaryText(props: IPersonaProps): JSX.Element {
      return (
        <div>
          <Icon iconName={ 'Suitcase' } className={ 'ms-JobIconExample' } />
          { props.secondaryText }
        </div>
      );
    }


    private async getMyPhotoDef(): Promise<string> {
        const msGraphToken = await this.getMSGraphAccessToken();
        let dfd: any = $.Deferred();

        var request = new XMLHttpRequest;
        request.open("GET", "https://graph.microsoft.com/v1.0/me/photo/$value");
        request.setRequestHeader("Authorization", "Bearer " + msGraphToken);
        request.responseType = "blob";
        request.onload = function () {
            if (request.readyState === 4 && request.status === 200) {
                var reader = new FileReader();
                reader.onload = function () {
                    dfd.resolve(reader.result);
                }
                reader.readAsDataURL(request.response);
            }
        };
        request.send(null);
        return dfd.promise();
    }

    private async getDetails(): Promise<IDisplayProps> {
        const msGraphToken = await this.getMSGraphAccessToken();
        let dfd: any = $.Deferred();

        let ajaxSettings: JQueryAjaxSettings = {
            url: "https://graph.microsoft.com/v1.0/me/",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + msGraphToken
            },
            success: (data) => {
                console.log(data);
                dfd.resolve(data);
            },
            error: (jqxr, errorCode, errorThrown) => {
                console.log(jqxr.responseText);
                dfd.reject("error");
            }
        }

        $.ajax(ajaxSettings);
        return dfd.promise();
    }

    private getMSGraphAccessToken(): Promise<string> {

        let dfd: any = $.Deferred();

        var requestHeaders: any = {
            'X-RequestDigest': $("#__REQUESTDIGEST").val(),
            "accept": "application/json;odata=nometadata",
            "content-type": "application/json;odata=nometadata"
        };
        var resourceData: any = {
            "resource": "https://graph.microsoft.com",
        };

        let ajaxSettings: JQueryAjaxSettings = {
            url: _spPageContextInfo.webAbsoluteUrl + "/_api/SP.OAuth.Token/Acquire",
            headers: requestHeaders,
            type: "POST",
            data: JSON.stringify(resourceData),
            success: (data) => {
                let msGraphToken: string = data.access_token;
                dfd.resolve(msGraphToken);
            },
            error: (jqxr, errorCode, errorThrown) => {
                console.log(jqxr.responseText);
                dfd.reject("error");
            }
        }
        $.ajax(ajaxSettings);
        return dfd.promise();
    }

}
