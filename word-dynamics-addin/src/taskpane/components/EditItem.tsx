/* eslint-disable prettier/prettier */
import * as React from "react";
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

export default class EditItem extends React.Component<EditItemProps> {

	onClick = async (notifyLoaded: () => void): Promise<void> => {
		// eslint-disable-next-line no-undef
		return Word.run(async (context) => {

			const doc = context.document;
			const originalRange = doc.getSelection();
			let displayName = SettingsService.getFieldDisplayName(this.props.entity, this.props.field, this.props.parentEntity);
			if (this.props.field.customFieldName) { // do a override of the fieldName, but after getting the fieldValue
				displayName = this.props.field.customFieldName;
			}
			let placeholder = originalRange.insertContentControl();
			placeholder.title = displayName;
			placeholder.insertText(displayName, "Replace");
			placeholder.tag = SettingsService.getFieldInternalName(this.props.entity, this.props.field, this.props.parentEntity);
			if (this.props.field.customFieldName) { // do a override of the fieldName, but after getting the fieldValue
				placeholder.tag = this.props.field.customFieldName;
			}

			await context.sync();

			// eslint-disable-next-line no-undef
			placeholder.getRange(Word.RangeLocation.after).select(Word.SelectionMode.end);
			await context.sync();

			// notify clear errors
			if (this.props.onError) {
				this.props.onError(null);
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
				if (this.props.onError) {
					this.props.onError(strings.failedAddingContentControls, error);
				}
				notifyLoaded();
			});
	}

	// eslint-disable-next-line no-undef
	render(): JSX.Element {
		return (<ClickableListItem
			onClick={this.onClick}
			showLoading={true}
			iconName={`ms-Icon ms-Icon--${this.props.type}`}
			label={this.props.field.displayName} />);
	}
}
