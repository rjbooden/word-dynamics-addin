/* eslint-disable prettier/prettier */
import { Entity, SettingsService } from "./SettingsService";
import { TokenService } from "./TokenService";

export interface IncludeData {
	entity: Entity;
	item: any;
}

export class DynamicsService {

	private static url = "/dataverse/api";

	public static getIncludedEntityData(entity: Entity, sourceItem: any): Promise<IncludeData[]> {
		if (entity.includeEntity) {
			let entityData = entity.includeEntity.map((entity) => {
				let includedEntity: Entity = SettingsService.getIncludedEntity(entity);
				if (includedEntity && sourceItem[entity.sourceField]) {
					return new Promise<IncludeData>((resolve, reject) => {
						let tempEntity: Entity = JSON.parse(JSON.stringify(includedEntity)); // deep copy
						tempEntity.additionalFilter = `${entity.targetField} eq '${sourceItem[entity.sourceField]}'`;
						tempEntity.topLimit = 1;
						this.getData(tempEntity, '').then((data) => {
							if (data && data.length) {
								resolve({
									entity: includedEntity,
									item: data[0]
								});
							}
							else {
								resolve({
									entity: includedEntity,
									item: {}
								});
							}
						}).catch((reason) => { reject(reason); });
					});
				}
				// return empty object
				return Promise.resolve({
					entity: includedEntity,
					item: {}
				});
			});
			return Promise.all(entityData);
		}
		return Promise.resolve(null);
	}

	public static getData(entity: Entity, query: string): Promise<any> {
		return new Promise<any>((resolve, reject) => {
			let fieldNames = entity.fields.map(field => {
				return field.fieldName;
			});
			if (!fieldNames.includes(entity.labelField)) {
				fieldNames.push(entity.labelField);
			}
			if (entity.includeEntity) {
				entity.includeEntity.forEach((entity) => {
					if (!fieldNames.includes(entity.sourceField)) {
						fieldNames.push(entity.sourceField);
					}
				});
			}
			query = query.replace("'", "''");
			let filter = !entity.additionalFilter ?
				`$filter=contains(${entity.searchField},'${query}')`
				: `$filter=(contains(${entity.searchField},'${query}') and (${entity.additionalFilter}))`;
			let topLimit = entity.topLimit ? `&$top=${entity.topLimit}` : "";
			let orderBy = entity.orderBy ? `&$orderby=${entity.orderBy}` : `&$orderby=${entity.labelField}`;
			let url = `${this.url}/${entity.tableName}?$select=${fieldNames.join(',')}&${filter}${topLimit}${orderBy}`;
			TokenService.AuthenticatedRequest<any>(url, {
				method: 'GET',
				cache: 'no-cache',
				headers: {
					'Content-Type': 'application/json',
				},
				redirect: 'follow',
				referrerPolicy: 'no-referrer'
			}).then((content) => {
				if (content.value && content.value.length) {
					resolve(content.value);
				}
				else {
					resolve([]);
				}
			}).catch((reason) => { reject(reason); });
		});
	}

}