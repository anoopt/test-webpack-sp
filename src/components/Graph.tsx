import * as React from "react";
import * as $ from "jquery";

export interface GraphProps { }
export interface IGraphState { 
    displayName: string;
    givenName: string;
    mobilePhone: string;
 }

export interface IDisplayProps {
    displayName: string;
    givenName: string;
    mobilePhone: string;
}

export class Graph extends React.Component<GraphProps, IGraphState> {

    constructor(props: GraphProps){
        super(props);
        this.state = {
            displayName: "",
            givenName: "",
            mobilePhone: ""
        };
    }

    public componentDidMount(): void {
        /* this.getName().then(r => {
            this.setState({
                yourName: r
            })
         }); */

         this.getDetails().then(r => {
            this.setState({
                displayName: r.displayName,
                givenName: r.givenName,
                mobilePhone: r.mobilePhone
            })
         });
         
      }

    public render(): React.ReactElement<GraphProps> {
        return (
        <div>
            <p>Your Name is {this.state.displayName}!</p>
            <p>Your Given Name is {this.state.givenName}!</p>
            <p>Your Mobile Number is {this.state.mobilePhone}!</p>
        </div>);
    }

    /* private async getName(): Promise<string> {
        const msGraphToken = await this.getMSGraphAccessToken();
        let dfd: any = $.Deferred();
        $.ajax({
            url: "https://graph.microsoft.com/v1.0/me/",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + msGraphToken
            },
            success: function (data) {
                console.log(data);
                let nameRetured: string = data.displayName;
                dfd.resolve(nameRetured);
            },
            error: function (jqxr, errorCode, errorThrown) {
                console.log(jqxr.responseText);
                dfd.reject("error");
            }
        });
        return dfd.promise();
    } */

    private async getDetails(): Promise<IDisplayProps> {
        const msGraphToken = await this.getMSGraphAccessToken();
        let dfd: any = $.Deferred();
        $.ajax({
            url: "https://graph.microsoft.com/v1.0/me/",
            headers: {
                "Accept": "application/json",
                "Authorization": "Bearer " + msGraphToken
            },
            success: function (data) {
                console.log(data);
                dfd.resolve(data);
            },
            error: function (jqxr, errorCode, errorThrown) {
                console.log(jqxr.responseText);
                dfd.reject("error");
            }
        });
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
        $.ajax({
            url: _spPageContextInfo.webAbsoluteUrl + "/_api/SP.OAuth.Token/Acquire",
            headers: requestHeaders,
            type: "POST",
            data: JSON.stringify(resourceData),
            success: function (data) {
                let msGraphToken: string = data.access_token;
                dfd.resolve(msGraphToken);
            },
            error: function (jqxr, errorCode, errorThrown) {
                console.log(jqxr.responseText);
                dfd.reject("error");
            }
        });
       return dfd.promise();
    }

}
