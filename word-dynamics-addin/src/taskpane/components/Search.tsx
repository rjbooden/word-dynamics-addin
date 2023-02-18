/* eslint-disable prettier/prettier */
import { SearchBox, Spinner, SpinnerSize } from "@fluentui/react";
import React, { useState, useEffect } from "react";
import { DynamicsService } from "../../services/DynamicsService";
import { strings } from "../../services/LocaleService";
import { Entity } from "../../services/SettingsService";
import ResultList from "./ResultList";

export interface SearchProps {
  selectedEntity: Entity;
  onError?: (errorMessage: string, errorInfo?: string) => void;
}

export default function Search(props: SearchProps) {

  const [searchValue, setSearchValue] = useState<string>('');
  const [resultItems, setResults] = useState<any[]>([]);
  const [isLoading, setLoading] = useState<boolean>(true);

  const onSearchBoxClear = () => {
    setSearchValue('');
    setResults([]);
  }

  const onSearchBoxChange = (_e, searchValue) => {
    setSearchValue(searchValue);
  }

  const startSearch = async (searchValue) => {
    setLoading(true);
    props.onError?.('');
    DynamicsService.getData(props.selectedEntity, searchValue).then(async (result) => {
      setLoading(false);
      setResults(result)
    }).catch((reason) => {
      // eslint-disable-next-line no-undef
      console.log(reason);
      setLoading(false);
      props.onError?.(strings.searchNotCompleted, reason);
    });
  }

  const onSearch = async (searchValue) => {
    startSearch(searchValue);
  }

  const autoSearch = () => {
    if (props.selectedEntity?.autoSearchEnabled) {
      startSearch('');
    }
    else {
      setLoading(false);
    }
  }

  useEffect(() => {
    setSearchValue('');
    setResults([]);
    // wait for updated state
    // eslint-disable-next-line no-undef
    setTimeout(autoSearch, 500);
  }, [props.selectedEntity]);

  return (
    <div>
      {
        <div>
          <span className="ms-font-m">
            {strings.searchFor} {props.selectedEntity?.displayName.toLowerCase()}:
          </span>
          <SearchBox
            placeholder={strings.search}
            value={searchValue}
            onSearch={onSearch}
            onChange={onSearchBoxChange}
            onClear={onSearchBoxClear} />
        </div>
      }
      {
        isLoading ?
          <Spinner size={SpinnerSize.large} className="loading-spinner" />
          :
          resultItems.length ?
            <ResultList
              message={strings.results}
              items={resultItems}
              onError={props.onError}
              entity={props.selectedEntity} />
            :
            null
      }
    </div>
  );

}
