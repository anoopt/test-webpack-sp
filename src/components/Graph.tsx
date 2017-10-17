import * as React from "react";
import * as $ from "jquery";

export interface GraphProps { }
export interface IGraphState { yourName: string }
export class Graph extends React.Component<GraphProps, IGraphState> {

    constructor(props: GraphProps){
        super(props);
        this.state = {
            yourName: ""
        };
    }

    public componentDidMount(): void {
        this.getName().then(r => {
            this.setState({
                yourName: r
            })
         });
         
      }

    public render(): React.ReactElement<GraphProps> {
        return <h1>Your name is {this.state.yourName}!</h1>;
    }

    private async getName(): Promise<any> {
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
    }


    private getMSGraphAccessToken(): Promise<any> {
        
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