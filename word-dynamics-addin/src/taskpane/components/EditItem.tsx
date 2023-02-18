/* eslint-disable prettier/prettier */
import React from "react";
import { strings } from "../../services/LocaleService";
import { Entity, EntityField, SettingsService } from "../../services/SettingsService";
import ClickableListItem from "./ClickableListItem";

export interface EditItemProps {
  type: string;
  parentEntity?: Entity;
  entity: Entity;
  field: EntityField;
  onError?: (errorMessage: string, errorInfo?: string) => void;
}

export default function EditItem(props: EditItemProps) {

  const onClick = async (notifyLoaded) => {
    // eslint-disable-next-line no-undef
    return Word.run(async (context) => {

      const doc = context.document;
      const originalRange = doc.getSelection();
      let displayName = SettingsService.getFieldDisplayName(props.entity, props.field, props.parentEntity);
      if (props.field.customFieldName) { // do a override of the fieldName, but after getting the fieldValue
        displayName = props.field.customFieldName;
      }
      let placeholder = originalRange.insertContentControl();
      placeholder.title = displayName;
      placeholder.insertText(displayName, "Replace");
      placeholder.tag = SettingsService.getFieldInternalName(props.entity, props.field, props.parentEntity);
      if (props.field.customFieldName) { // do a override of the fieldName, but after getting the fieldValue
        placeholder.tag = props.field.customFieldName;
      }

      await context.sync();

      // eslint-disable-next-line no-undef
      placeholder.getRange(Word.RangeLocation.after).select(Word.SelectionMode.end);
      await context.sync();

      // notify clear errors
      if (props.onError) {
        props.onError(null);
      }
      notifyLoaded();
    })
      .catch(function (error) {
        // eslint-disable-next-line no-undef
        console.log("Error: " + error);
        // eslint-disable-next-line no-undef
        if (error instanceof OfficeExtension.Error) {
          // eslint-disable-next-line no-undef
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
        if (props.onError) {
          props.onError(strings.failedAddingContentControls, error);
        }
        notifyLoaded();
      });
  }

  return (<ClickableListItem
    onClick={onClick}
    showLoading={true}
    iconName={`ms-Icon ms-Icon--${props.type}`}
    label={props.field.displayName} />);
}
