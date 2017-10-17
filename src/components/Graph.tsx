import * as React from "react";
import * as $ from "jquery";

interface GraphProps { }
interface IGraphState {
    myData: IDisplayProps
}

interface IDisplayProps {
    displayName: string;
    givenName: string;
    mobilePhone: string;
    imageUrl: string;
}

export class Graph extends React.Component<GraphProps, IGraphState> {

    constructor(props: GraphProps) {
        super(props);
        this.state = {
            myData: {
                displayName: "",
                givenName: "",
                mobilePhone: "",
                imageUrl: ""
            }
        };
    }

    public componentDidMount(): void {
        /* this.getName().then(r => {
            this.setState({
                yourName: r
            })
         }); */

        this.getDetails().then(r => {
            this.getMyPhotoDef().then(i => {
                this.setState({
                    myData: {
                        displayName: r.displayName,
                        givenName: r.givenName,
                        mobilePhone: r.mobilePhone,
                        imageUrl: i
                    }
                })
            })
        });

        this.getMyPhoto();

    }

    public render(): React.ReactElement<GraphProps> {
        return (
            <div>
                <p>Your name is {this.state.myData.displayName},</p>
                <p>Your given name is {this.state.myData.givenName},</p>
                <p>Your mobile number is {this.state.myData.mobilePhone}</p>
                <img src={this.state.myData.imageUrl} />
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

    private async getMyPhoto() {
        const msGraphToken = await this.getMSGraphAccessToken();

        var request = new XMLHttpRequest;
        request.open("GET", "https://graph.microsoft.com/v1.0/me/Photos/48X48/$value");
        request.setRequestHeader("Authorization", "Bearer " + msGraphToken);
        request.responseType = "blob";
        request.onload = function () {
            if (request.readyState === 4 && request.status === 200) {
                var imageElm = document.createElement("img");
                var reader = new FileReader();
                reader.onload = function () {
                    // Add the base64 image to the src attribute
                    imageElm.src = reader.result;
                    // Display the user's profile picture
                    document.getElementsByClassName('user-picture-box')[0].appendChild(imageElm);
                }
                reader.readAsDataURL(request.response);
            }
        };
        request.send(null);
    }

    private async getMyPhotoDef(): Promise<string> {
        const msGraphToken = await this.getMSGraphAccessToken();
        let dfd: any = $.Deferred();

        var request = new XMLHttpRequest;
        request.open("GET", "https://graph.microsoft.com/v1.0/me/Photos/48X48/$value");
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
