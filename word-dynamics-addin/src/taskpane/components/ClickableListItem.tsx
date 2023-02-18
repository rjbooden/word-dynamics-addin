/* eslint-disable prettier/prettier */
import { Spinner, SpinnerSize } from "@fluentui/react";
import * as React from "react";

export interface ClickableListItemProps {
	label: string;
	iconName?: string;
	showLoading?: boolean;
	onClick: (notifyLoaded?: () => void) => void;
	onDelete?: () => void;
}

export interface ClickableListItemState {
	isLoading: boolean;
}

export default class ClickableListItem extends React.Component<ClickableListItemProps, ClickableListItemState> {

	constructor(props, state) {
		super(props, state);
		this.state = { isLoading: false };
	}

	notifyLoaded = () => {
		if (this.props.showLoading) {
			this.setState({ isLoading: false });
		}
	}

	onClick = () => {
		if (this.props.showLoading) {
			this.setState({ isLoading: true });
		}
		this.props.onClick(this.notifyLoaded);
	}

	render() {
		const { label, iconName } = this.props;

		return (
			<li className="ms-ListItem clickable" onClick={this.onClick}>
				{iconName && !this.state.isLoading ? <i className={`ms-Icon ms-Icon--${iconName}`}></i> : null}
				{this.state.isLoading ? <Spinner size={SpinnerSize.xSmall} className="item-loading-spinner" /> : null}
				<span className="ms-font-m ms-fontColor-neutralPrimary item-label">{label}</span>
				{this.props.onDelete ?
					<i className={`ms-Icon ms-Icon--Delete item-delete`}
						onClick={(e) => {
							e.stopPropagation();
							this.props.onDelete();
							this.setState({ isLoading: true });
						}}></i>
					: null}
			</li>
		);
	}
}
