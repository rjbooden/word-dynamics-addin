/* eslint-disable prettier/prettier */
import { SearchBox, Spinner, SpinnerSize } from "@fluentui/react";
import * as React from "react";
import { DynamicsService } from "../../services/DynamicsService";
import { strings } from "../../services/LocaleService";
import { Entity } from "../../services/SettingsService";
import ResultList from "./ResultList";

export interface SearchProps {
  selectedEntity: Entity;
  onError?: (errorMessage: string, errorInfo?: string) => void;
}

export interface SearchState {
  isLoading: boolean;
  resultItems: any[];
  searchValue: string;
}

export default class Search extends React.Component<SearchProps, SearchState> {

  constructor(props, context) {
    super(props, context);
    this.state = {
      searchValue: '',
      resultItems: [],
      isLoading: true
    };
  }

  onSearchBoxClear = () => {
    this.setState({ searchValue: '', resultItems: [] });
  }

  onSearchBoxChange = (_e, searchValue) => {
    this.setState({ searchValue });
  }

  startSearch = async (searchValue) => {
    this.setState({ isLoading: true, searchValue });
    this.props.onError?.('');
    DynamicsService.getData(this.props.selectedEntity, searchValue).then(async (result) => {
      this.setState({ resultItems: result, isLoading: false });
    }).catch((reason) => {
      // eslint-disable-next-line no-undef
      console.log(reason);
      this.setState({ isLoading: false });
      this.props.onError?.(strings.searchNotCompleted, reason);
    });
  }

  onSearch = async (searchValue) => {
    this.startSearch(searchValue);
  }

  autoSearch = () => {
    if (this.props.selectedEntity?.autoSearchEnabled) {
      this.startSearch('');
    }
    else {
      this.setState({ isLoading: false });
    }
  }

  componentDidMount(): void {
    this.autoSearch();
  }

  componentDidUpdate(prevProps) {
    if (prevProps.selectedEntity !== this.props.selectedEntity) {
      this.setState({ searchValue: '', resultItems: [] });
      // wait for updated state
      // eslint-disable-next-line no-undef
      setTimeout(this.autoSearch, 500);
    }
  }

  render() {
    return (
      <div>
        {
          <div>
            <span className="ms-font-m">
              {strings.searchFor} {this.props.selectedEntity?.displayName.toLowerCase()}:
            </span>
            <SearchBox
              placeholder={strings.search}
              value={this.state.searchValue}
              onSearch={this.onSearch}
              onChange={this.onSearchBoxChange}
              onClear={this.onSearchBoxClear} />
          </div>
        }
        {
          this.state.isLoading ?
            <Spinner size={SpinnerSize.large} className="loading-spinner" />
            :
            this.state.resultItems.length ?
              <ResultList message={strings.results}
                items={this.state.resultItems}
                onError={this.props.onError}
                entity={this.props.selectedEntity} />
              :
              null
        }
      </div>
    );
  }
}
