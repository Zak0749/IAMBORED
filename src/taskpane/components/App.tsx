/* eslint-disable @typescript-eslint/no-unused-vars */
import React, { useState } from "react";
import { ComboBox, DefaultButton, TextField, IComboBoxOption, Stack, IStackTokens } from "@fluentui/react";
import axios from "axios";

export interface AppProps {
  subject: string;
  attendees: { name: string; email: string }[];
}

const containerStackTokens: IStackTokens = { childrenGap: 40 };

export function App(props: AppProps) {
  let [options, setOptions] = useState<IComboBoxOption[]>([]);
  let [extraInfo, setExtraInfo] = useState("");
  let [selected, setSelected] = useState("");
  let [message, setMessage] = useState("");

  let upload = async () => {
    try {
      let object = {
        subject: props.subject,
        attendees: props.attendees,
        extraInfo: extraInfo,
        selected: selected,
      };

      let response = await axios.post("https://alphanumericadvancedkeyboardmapping.zak0749.repl.co/upload/", object);

      if (response.status == 200) {
        setMessage("sucess");
      } else {
        setMessage("There was an error");
      }
    } catch (err) {
      setMessage("There was an error");
    }
  };

  const updateOptions = async (_option: IComboBoxOption, _index: number, value: string) => {
    if (!value) {
      return;
    }
    let response = await axios("https://alphanumericadvancedkeyboardmapping.zak0749.repl.co/matches/" + value);

    setOptions(
      response.data.map((val) => {
        return {
          key: val,
          text: val,
        };
      })
    );
  };

  return (
    <Stack tokens={containerStackTokens} className="main">
      <TextField label="Additional Information" multiline rows={3} onChange={(_, value) => setExtraInfo(value)} />
      <ComboBox
        placeholder="Select a letter code"
        onPendingValueChanged={updateOptions}
        onChange={(_a, _b, _c, value) => setSelected(value)}
        options={options}
        allowFreeInput
        autoComplete="on"
      />
      <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={upload}>
        Sumbit
      </DefaultButton>
      <div>{message}</div>
    </Stack>
  );
}
