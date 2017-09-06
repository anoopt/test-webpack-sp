import * as React from "react";
import pnp from "sp-pnp-js";

export interface HelloProps { }
export interface IHelloState { webTitle: string }

// 'HelloProps' describes the shape of props.
// State is never set so we use the 'undefined' type.
export class Hello extends React.Component<HelloProps, IHelloState> {

    constructor(props: HelloProps){
        super(props);
        this.state = {
          webTitle: ""
        };
    }

    public componentDidMount(): void {
        this.getWebTitle();
      }

    public render(): React.ReactElement<HelloProps> {
        return <h1>Web Title {this.state.webTitle}!</h1>;
    }

    private async getWebTitle(): Promise<void>{
        pnp.sp.web.get().then(r => {
           this.setState({
               webTitle: r.Title
           })
        });
    }
}