/* eslint-disable prettier/prettier */
import * as React from "react";
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

export interface AppState {
  isLoading: boolean;
  selectedEntity?: Entity;
  modeToggleIcon: string;
  errorMessage?: string;
  errorInfo?: string;
}

export default class App extends React.Component<AppProps, AppState> {

  constructor(props, context) {
    super(props, context);
    this.state = {
      selectedEntity: null,
      isLoading: true,
      modeToggleIcon: SettingsService.getMode()
    };
  }

  componentDidMount() {
    SettingsService.getSettings().then(() => {
      // auto select first entity by default
      if (SettingsService.entities.length === 0) {
        // todo: display error message
        this.setState({ errorMessage: strings.unableToLoadSettings, isLoading: false });
        return;
      }
      let selectedEntity: Entity = SettingsService.getInitialEntity();
      this.setState({ selectedEntity, isLoading: false });
    }).catch((reason) => {
      // eslint-disable-next-line no-undef
      console.log(reason);
      this.setState({
        errorMessage: strings.unableToLoadSettings,
        errorInfo: reason,
        isLoading: false
      });
    });
  }

  modeButtonClick = (_e, item: IContextualMenuItem) => {
    SettingsService.setMode(item.key as Mode);
    this.setState({ modeToggleIcon: item.key });
  }

  onEntityChange = (_e, value) => {
    let selectedEntity = SettingsService.entities.find((s) => { return s.uniqueName == value.key; });
    SettingsService.setInitialEnity(selectedEntity);
    this.setState({ selectedEntity });
  }

  onError = (errorMessage: string, errorInfo?: string) => {
    this.setState({ errorMessage, errorInfo });
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

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
        onClick: this.modeButtonClick
      });
    }

    return (
      <div className="ms-welcome">
        <Header logo="/assets/icon-64.png" title={this.props.title} message={strings.title} />
        {
          this.state.errorMessage ?
            <ErrorBox message={this.state.errorMessage} info={this.state.errorInfo} />
            : null
        }
        {
          entityOptions.length ?
            <div className="full-width-24">
              <ComboBox
                defaultSelectedKey={this.state.selectedEntity.uniqueName}
                options={entityOptions}
                calloutProps={{ doNotLayer: true }}
                className="width-minus-24"
                onChange={this.onEntityChange}
                disabled={this.state.modeToggleIcon === Mode.Copy}
              />
              <IconButton
                className="align-right"
                iconProps={{ iconName: this.state.modeToggleIcon }}
                menuProps={menuProps}
                onRenderMenuIcon={() => null} />
            </div>
            :
            null
        }
        {
          this.state.selectedEntity ?
            <div className="default-container">
              {
                this.state.modeToggleIcon === Mode.Search ?
                  <Search selectedEntity={this.state.selectedEntity} onError={this.onError}></Search>
                  : this.state.modeToggleIcon === Mode.Edit ?
                    <EditList message={strings.clickFieldToAdd} onError={this.onError} entity={this.state.selectedEntity} />
                    : this.state.modeToggleIcon === Mode.Copy ?
                      <SnippetsView onError={this.onError}></SnippetsView>
                      : null
              }
            </div>
            : null
        }
        {
          this.state.isLoading ?
            <Spinner size={SpinnerSize.large} className="loading-spinner" />
            : null
        }
      </div>
    );
  }
}
