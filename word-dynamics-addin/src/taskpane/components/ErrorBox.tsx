import * as React from "react";

export interface ErrorProps {
  message: string;
  info: string;
}

export default class ErrorBox extends React.Component<ErrorProps> {
  render() {
    const { message, info } = this.props;

    return (
      <section className="error-pane ms-bgColor-neutralLighter ms-u-fadeIn500" title={info}>
        <h2 className="ms-fontSize-m ms-fontWeight-bold ms-fontColor-alert">{message}</h2>
      </section>
    );
  }
}
