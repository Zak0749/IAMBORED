import React, { useState } from "react";
import { DefaultButton } from "@fluentui/react";
import axios from "axios";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  person: string;
  message: string;
}

export function App(props: AppProps) {
  let [person, setPerson] = useState("No One");
  let [message, setMessage] = useState("Asked about outlook");

  let onclick = async () => {
    try {
      let response = await axios("https://alphanumericadvancedkeyboardmapping.zak0749.repl.co");
      let data = response.data as AppState;

      setPerson(data.person);
      setMessage(data.message);
    } catch (error) {
      setMessage("error:" + error);
    }
  };

  return (
    <div className="ms-welcome">
      <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={onclick}>
        Run
      </DefaultButton>

      <h1>{person}</h1>

      <p>{message}</p>

      {props.isOfficeInitialized}

      {props.title}
    </div>
  );
}

// export default class App extends React.Component<AppProps, AppState> {
//   constructor(props, context) {
//     super(props, context);
//     this.state = {
//       person: "No One",
//       message: "Asked about outlook",
//     };
//   }

//   componentDidMount() {}

//   click = async () => {};

//   render() {
//     return (
//       <div className="ms-welcome">
//         <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
//           Run
//         </DefaultButton>

//         <h1>{this.state.person}</h1>

//         <p>{this.state.message}</p>
//       </div>
//     );
//   }
// }
