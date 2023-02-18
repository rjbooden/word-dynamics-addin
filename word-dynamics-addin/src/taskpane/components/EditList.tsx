/* eslint-disable prettier/prettier */
import React from "react";
import { Entity } from "../../services/SettingsService";
import EditItem from "./EditItem";
import EditSubList from "./EditSubList";

export interface EditListProps {
  message: string;
  entity: Entity;
  onError?: (errorMessage: string, errorInfo?: string) => void;
}

export default function EditList(props: EditListProps) {

  const { message, entity } = props;

  const listItems = entity.fields.map((field, index) => (
    <EditItem type="TextField" field={field} entity={props.entity} onError={props.onError} key={index} />
  ));

  let includeEntityItems = null;
  if (entity.includeEntity) {
    includeEntityItems = entity.includeEntity.map((entity, index) => (
      <EditSubList entity={entity} parentEntity={props.entity} onError={props.onError} key={index} />
    ));
  }

  return (
    <main className="ms-welcome__main">
      <h2 className="ms-font-m ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">{message}</h2>
      <ul className="ms-List ms-welcome__features ms-u-slideUpIn10">
        {listItems}
        {includeEntityItems}
      </ul>
    </main>
  );
}
