/* eslint-disable prettier/prettier */
import { Spinner, SpinnerSize } from "@fluentui/react";
import * as React from "react";

export interface ClickableListItemProps {
	label: string;
	iconName?: string;
	showLoading?: boolean;
	onClick?: (notifyLoaded?: () => void) => void;
	onDelete?: () => void;
}

export interface ClickableListItemState {
	isLoading: boolean;
}

export default class ClickableListItem extends React.Component<ClickableListItemProps, ClickableListItemState> {

	constructor(props: ClickableListItemProps, state: ClickableListItemState) {
		super(props, state);
		this.state = { isLoading: false };
	}

	notifyLoaded = (): void => {
		if (this.props.showLoading) {
			this.setState({ isLoading: false });
		}
	}

	onClick = (): void => {
		if (this.props.onClick) {
			if (this.props.showLoading) {
				this.setState({ isLoading: true });
			}
			this.props.onClick(this.notifyLoaded);
		}
	}

	// eslint-disable-next-line no-undef
	render(): JSX.Element {
		const { label, iconName } = this.props;

		return (
			<li className={`ms-ListItem ${this.props.onClick ? 'clickable' : ''}`}  onClick={this.onClick}>
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
