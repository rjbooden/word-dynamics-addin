/* eslint-disable prettier/prettier */
import * as React from "react";
import { IncludeEntity, Entity, SettingsService } from "../../services/SettingsService";
import EditItem from "./EditItem";

export interface EditSubListProps {
  entity: IncludeEntity;
  parentEntity: Entity;
  onError?: (errorMessage: string, errorInfo?: string) => void;
}

export default class EditSubList extends React.Component<EditSubListProps> {
  render() {
    const { entity, parentEntity } = this.props;

    let includedEntity: Entity = SettingsService.getIncludedEntity(entity);
    const listItems = includedEntity.fields.map((field, index) => (
      <EditItem type="TextField" field={field} entity={includedEntity} parentEntity={parentEntity} onError={this.props.onError} key={index} />
    ));

    return (
      <li>
        {includedEntity.displayName}:
        <ul className="ms-List nextLevel ms-welcome__features ms-u-slideUpIn10">
          {listItems}
        </ul>
      </li>
    );
  }
}
