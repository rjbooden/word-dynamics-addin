/* eslint-disable prettier/prettier */

import { ContentSource, TokenService } from "./TokenService";

// Modes should be Fluent UI Core Icons names
export enum Mode {
	Search = "Search",
	Edit = "Edit",
	Copy = "Copy"
}

export interface Snippet {
	title: string;
	content: string;
}

export interface EntityField {
	displayName: string;
	fieldName: string;
	customFieldName?: string;
	containsHtml?: boolean;
}

export interface IncludeEntity {
	uniqueName: string;
	sourceField: string;
	targetField: string;
}

export interface Entity {
	uniqueName: string;
	displayName: string;
	tableName: string;
	isHidden?: boolean;
	labelField: string;
	fields: EntityField[];
	searchField: string;
	additionalFilter?: string;
	autoSearchEnabled?: boolean;
	iconName: string;
	includeEntity?: IncludeEntity[];
	topLimit?: number;
	orderBy?: string;
}

export class SettingsService {
	public static entities: Entity[] = [];
	public static Snippets: Snippet[] = [];
	static selectedEntity: string = 'wdai_selected_entity';
	static selectedMode: string = 'wdai_selected_mode';

	public static setInitialEnity(entity: Entity) {
		try {
			// eslint-disable-next-line no-undef
			localStorage.setItem(SettingsService.selectedEntity, entity.uniqueName);
		}
		// eslint-disable-next-line no-undef
		catch (e) { console.log(e); } // on exception, log and continue
	}

	public static getInitialEntity(): Entity {
		try {
			// eslint-disable-next-line no-undef
			let selected = localStorage.getItem(SettingsService.selectedEntity);
			if (selected) {
				let found = this.entities.find((s) => { return s.uniqueName === selected; });
				if (found) {
					return found;
				}
			}
		}
		// eslint-disable-next-line no-undef
		catch (e) { console.log(e); } // on exception, log and continue
		return this.entities.find((s) => { return !s.isHidden; });
	}

	public static setMode(mode: Mode) {
		try {
			// eslint-disable-next-line no-undef
			localStorage.setItem(SettingsService.selectedMode, mode);
		}
		// eslint-disable-next-line no-undef
		catch (e) { console.log(e); } // on exception, log and continue
	}

	public static getMode(): Mode {
		try {
			// eslint-disable-next-line no-undef
			let selected = localStorage.getItem(SettingsService.selectedMode) as Mode;
			if (selected) {
				return selected;
			}
		}
		// eslint-disable-next-line no-undef
		catch (e) { console.log(e); } // on exception, log and continue
		return Mode.Search;
	}

	public static getIncludedEntity(entity: IncludeEntity): Entity {
		return this.entities.find((s) => { return s.uniqueName === entity.uniqueName; });
	}

	public static getFieldDisplayName(entity: Entity, field: EntityField, parentEntity: Entity = null): string {
		if (parentEntity) {
			return `${parentEntity.displayName}:${entity.displayName}:${field.displayName}`;
		}
		return `${entity.displayName}:${field.displayName}`;
	}

	public static getFieldInternalName(entity: Entity, field: EntityField, parentEntity: Entity = null): string {
		if (parentEntity) {
			return `${parentEntity.uniqueName}:${entity.uniqueName}:${field.fieldName}`;
		}
		return `${entity.uniqueName}:${field.fieldName}`;
	}

	public static saveSnippets(snippets: Snippet[]): Promise<void> {
		return new Promise<void>((resolve, reject) => {
			TokenService.AuthenticatedRequest<void>("/Snippets", {
				body: JSON.stringify(snippets),
				method: 'POST',
				cache: 'no-cache',
				headers: {
					'Content-Type': 'application/json',
				},
				redirect: 'follow',
				referrerPolicy: 'no-referrer'
			}, ContentSource.None).then(() => {
				resolve();
			}).catch((reason) => { reject(reason); });
		});
	}

	public static loadSnippets(): Promise<Snippet[]> {
		return new Promise<Snippet[]>((resolve, reject) => {
			TokenService.AuthenticatedRequest<Snippet[]>("/Snippets", {
				method: 'GET',
				cache: 'no-cache',
				headers: {
					'Content-Type': 'application/json'
				},
				redirect: 'follow',
				referrerPolicy: 'no-referrer'
			}).then((snippets) => {
				this.Snippets = snippets;
				resolve(this.Snippets);
			}).catch((reason) => { reject(reason); });
		});
	}

	public static async getSettings(): Promise<void> {
		return new Promise<void>((resolve, reject) => {
			// eslint-disable-next-line no-undef
			fetch("/settings.json").then((response) => {
				if (response.status == 200) {
					response.json().then((jsonObject) => {
						this.entities = jsonObject;
						resolve();
					}).catch((reason) => {
						reject(reason);
					});
				}
				else {
					reject(`${response.status} : ${response.statusText}`);
				}
			}).catch((reason) => {
				reject(reason);
			});
		});
	}
}