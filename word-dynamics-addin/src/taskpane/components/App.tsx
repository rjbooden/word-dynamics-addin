/* eslint-disable prettier/prettier */
import React, { useState, useEffect } from "react";
import Header from "./Header";
import Progress from "./Progress";
import { SettingsService, Entity, Mode } from "../../services/SettingsService";
import { ComboBox, IComboBoxOption, IconButton, IContextualMenuProps, IContextualMenuItem, Spinner, SpinnerSize } from "@fluentui/react";
import EditList from "./EditList";
import ErrorBox from "./ErrorBox";
import Search from "./Search";
import { strings } from "../../services/LocaleService";
import SnippetsView from "./SnippetsView";

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global Word, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export default function App(props: AppProps) {

  const [selectedEntity, setEntity] = useState<Entity | null>(null);
  const [isLoading, setLoading] = useState<boolean>(true);
  const [modeToggleIcon, setMode] = useState<Mode>(SettingsService.getMode());
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [errorInfo, setErrorInfo] = useState<string | null>(null);

  useEffect(() => {
    SettingsService.getSettings().then(() => {
      // auto select first entity by default
      if (SettingsService.entities.length === 0) {
        // todo: display error message
        setErrorMessage(strings.unableToLoadSettings);
        setLoading(false);
        return;
      }
      let selectedEntity: Entity = SettingsService.getInitialEntity();
      setEntity(selectedEntity);
      setLoading(false);
    }).catch((reason) => {
      // eslint-disable-next-line no-undef
      console.log(reason);
      setErrorMessage(strings.unableToLoadSettings);
      setErrorInfo(reason);
      setLoading(false);
    });
  }, []);

  const modeButtonClick = (_e, item: IContextualMenuItem) => {
    SettingsService.setMode(item.key as Mode);
    setMode(item.key as Mode);
  }

  const onEntityChange = (_e, value) => {
    let selectedEntity = SettingsService.entities.find((s) => { return s.uniqueName == value.key; });
    SettingsService.setInitialEnity(selectedEntity);
    setEntity(selectedEntity);
  }

  const onError = (errorMessage: string, errorInfo?: string) => {
    setErrorMessage(errorMessage);
    setErrorInfo(errorInfo);
  }

  const { title, isOfficeInitialized } = props;

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message={strings.pleaseSideload}
      />
    );
  }

  const entityOptions: IComboBoxOption[] = [];
  SettingsService.entities?.forEach((entity) => {
    if (!entity.isHidden) {
      entityOptions.push({ key: entity.uniqueName, text: entity.displayName });
    }
  });

  const menuProps: IContextualMenuProps = {
    items: [],
    directionalHintFixed: true,
    calloutProps: {
      calloutWidth: 36
    }
  };

  for (let item in Mode) {
    menuProps.items.push({
      key: item,
      iconProps: { iconName: item },
      onClick: modeButtonClick
    });
  }

  return (
    <div className="ms-welcome">
      <Header logo="/assets/icon-64.png" title={props.title} message={strings.title} />
      {
        errorMessage ?
          <ErrorBox message={errorMessage} info={errorInfo} />
          : null
      }
      {
        entityOptions.length ?
          <div className="full-width-24">
            <ComboBox
              defaultSelectedKey={selectedEntity.uniqueName}
              options={entityOptions}
              calloutProps={{ doNotLayer: true }}
              className="width-minus-24"
              onChange={onEntityChange}
              disabled={modeToggleIcon === Mode.Copy}
            />
            <IconButton
              className="align-right"
              iconProps={{ iconName: modeToggleIcon }}
              menuProps={menuProps}
              onRenderMenuIcon={() => null} />
          </div>
          :
          null
      }
      {
        selectedEntity ?
          <div className="default-container">
            {
              modeToggleIcon === Mode.Search ?
                <Search selectedEntity={selectedEntity} onError={onError}></Search>
                : modeToggleIcon === Mode.Edit ?
                  <EditList message={strings.clickFieldToAdd} onError={onError} entity={selectedEntity} />
                  : modeToggleIcon === Mode.Copy ?
                    <SnippetsView onError={onError}></SnippetsView>
                    : null
            }
          </div>
          : null
      }
      {
        isLoading ?
          <Spinner size={SpinnerSize.large} className="loading-spinner" />
          : null
      }
    </div>
  );
}

