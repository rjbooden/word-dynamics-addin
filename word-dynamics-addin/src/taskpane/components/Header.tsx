/* eslint-disable prettier/prettier */
import React from "react";

export interface HeaderProps {
  title: string;
  logo: string;
  message: string;
}

export default function Header(props: HeaderProps) {
  const { title, logo, message } = props;

  return (
    <section className="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
      {
        (logo == null) ? null :
          <img width="50" height="50" src={logo} alt={title} title={title} />
      }
      <h1 className="ms-fontSize-l ms-fontWeight-bold ms-fontColor-neutralPrimary">{message}</h1>
    </section>
  );

}
