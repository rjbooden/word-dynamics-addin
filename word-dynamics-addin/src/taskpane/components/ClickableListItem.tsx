/* eslint-disable prettier/prettier */
import { Spinner, SpinnerSize } from "@fluentui/react";
import React, { useState } from "react";

export interface ClickableListItemProps {
  label: string;
  iconName?: string;
  showLoading?: boolean;
  onClick: (notifyLoaded?: () => void) => void;
  onDelete?: () => void;
}

export default function ClickableListItem(props: ClickableListItemProps) {
  const [isLoading, setLoading] = useState(false);

  const notifyLoaded = () => {
    if (props.showLoading) {
      setLoading(false);
    }
  }

  const onClick = () => {
    if (props.showLoading) {
      setLoading(true);
    }
    props.onClick(notifyLoaded);
  }

  const { label, iconName } = props;

  return (
    <li className="ms-ListItem clickable" onClick={onClick}>
      {iconName && !isLoading ? <i className={`ms-Icon ms-Icon--${iconName}`}></i> : null}
      {isLoading ? <Spinner size={SpinnerSize.xSmall} className="item-loading-spinner" /> : null}
      <span className="ms-font-m ms-fontColor-neutralPrimary item-label">{label}</span>
      {props.onDelete ?
        <i className={`ms-Icon ms-Icon--Delete item-delete`}
          onClick={(e) => {
            e.stopPropagation();
            props.onDelete();
            setLoading(true);
          }}></i>
        : null}
    </li>
  );
}
