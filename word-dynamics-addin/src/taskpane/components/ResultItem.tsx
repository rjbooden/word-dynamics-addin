/* eslint-disable prettier/prettier */
import React from "react";
import { DynamicsService } from "../../services/DynamicsService";
import { strings } from "../../services/LocaleService";
import { Entity, SettingsService } from "../../services/SettingsService";
import ClickableListItem from "./ClickableListItem";

export interface ResultItemProps {
  item: any;
  entity?: Entity;
  onError?: (errorMessage: string, errorInfo?: string) => void;
}

export default function ResultItem(props: ResultItemProps) {

  const onClick = async (notifyLoaded) => {
    await fillFields(props.entity, props.item);
    DynamicsService.getIncludedEntityData(props.entity, props.item).then((includeDataResults) => {
      if (includeDataResults) {
        includeDataResults.forEach(async (includeData) => {
          await fillFields(includeData.entity, includeData.item, props.entity);
        });
      }
      notifyLoaded();
    }).catch((reason) => {
      // eslint-disable-next-line no-undef
      console.log(reason);
      if (props.onError) {
        props.onError(strings.failedRetrievingIncludedEntity, reason);
        notifyLoaded();
      }
    });
  }

	const getFieldValue = (fieldName:string, item: any) => {
		if (fieldName.indexOf(',') < 0) {
			return item[fieldName] || " ";
		}
		else {
			var values = fieldName.split(',').map((name) => {
				return item[name]; 
			});
			var fieldValue = values.filter(entry => /\S/.test(entry)).join(' ').trim();
			if (fieldValue.length == 0) {
				fieldValue = " ";
			}
			return fieldValue;
		}
	}

  const fillFields = async (entity: Entity, item: any, parentEntity: Entity = null) => {
    for (var i = 0; i < entity.fields.length; i++) {
      let field = entity.fields[i];
      let fieldName = SettingsService.getFieldInternalName(entity, field, parentEntity);
      let fieldValue = getFieldValue(field.fieldName, item);
      if (field.customFieldName) { // do a override of the fieldName, but after getting the fieldValue
        fieldName = field.customFieldName;
      }

      // eslint-disable-next-line no-undef
      await Word.run(async (context) => {
        const doc = context.document;

        var placeholder = doc.contentControls.getByTag(fieldName);
        // eslint-disable-next-line office-addins/no-navigational-load
        placeholder.load("items");
        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();

        // eslint-disable-next-line office-addins/load-object-before-read
        for (var i = 0; i < placeholder.items.length; i++) {
          // eslint-disable-next-line office-addins/load-object-before-read
          var item = placeholder.items[i];
          if (field.containsHtml) {
            if (fieldValue === ' ') {
              fieldValue = '&nbsp;';
            }
            item.insertHtml(fieldValue, "Replace");
          }
          else {
            item.insertText(fieldValue, "Replace");
          }
        }

        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();

        // notify clear errors
        if (props.onError) {
          props.onError(null);
        }
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
          props.onError(strings.failedSettingContentControlValues, error);
        }
      });
    }
  }

  return (<ClickableListItem
    onClick={onClick}
    showLoading={true}
    iconName={props.entity.iconName}
    label={getFieldValue(props.entity.labelField, props.item)} />);
}
