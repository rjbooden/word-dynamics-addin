import * as React from "react";
import { Entity } from "../../services/SettingsService";
import ResultItem from "./ResultItem";

export interface ResultListProps {
  message: string;
  items: any[];
  entity?: Entity;
  onError?: (errorMessage: string, errorInfo?: string) => void;
}

export default class ResultList extends React.Component<ResultListProps> {
  render() {
    const { items, message, entity } = this.props;

    const listItems = items.map((item, index) => (
      <ResultItem item={item} entity={entity} onError={this.props.onError} key={index} />
    ));
    return (
      <main className="ms-welcome__main">
        <h2 className="ms-font-l ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{message}</h2>
        <ul className="ms-List ms-welcome__features ms-u-slideUpIn10">{listItems}</ul>
      </main>
    );
  }
}
