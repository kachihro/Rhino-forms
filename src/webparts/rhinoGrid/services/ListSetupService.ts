import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import { IGridConfig, IFieldDefinition } from '../types/IFieldConfig';

export class ListSetupService {
  constructor(private sp: SPFI) {}

  public async listExists(listName: string): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(listName).select('Id')();
      return true;
    } catch {
      return false;
    }
  }

  public async provisionList(config: IGridConfig): Promise<void> {
    const exists = await this.listExists(config.listName);

    if (!exists) {
      await this.sp.web.lists.add(config.listName, config.listDescription || '', 100, false);
    }

    const existingFields = await this.sp.web.lists
      .getByTitle(config.listName)
      .fields
      .filter('Hidden eq false')
      .select('InternalName')();

    const existingNames = new Set(existingFields.map((f: { InternalName: string }) => f.InternalName.toLowerCase()));

    for (const field of config.fields) {
      if (field.internalName.toLowerCase() === 'title') {
        if (field.displayName && field.displayName !== 'Title') {
          await this.sp.web.lists
            .getByTitle(config.listName)
            .fields
            .getByInternalNameOrTitle('Title')
            .update({ Title: field.displayName });
        }
        continue;
      }

      if (existingNames.has(field.internalName.toLowerCase())) {
        continue;
      }

      await this.addFieldAsXml(config.listName, field);
    }
  }

  private async addFieldAsXml(listName: string, field: IFieldDefinition): Promise<void> {
    const required = field.required ? 'TRUE' : 'FALSE';
    const displayName = this.escapeXml(field.displayName);
    const name = this.escapeXml(field.internalName);
    let schemaXml = '';

    switch (field.type) {
      case 'Text':
        schemaXml = `<Field Type="Text" Name="${name}" StaticName="${name}" DisplayName="${displayName}" MaxLength="${field.maxLength || 255}" Required="${required}"/>`;
        break;
      case 'Note':
        schemaXml = `<Field Type="Note" Name="${name}" StaticName="${name}" DisplayName="${displayName}" NumLines="6" RichText="FALSE" Required="${required}"/>`;
        break;
      case 'Choice': {
        const choices = (field.choices || []).map(c => `<CHOICE>${this.escapeXml(c)}</CHOICE>`).join('');
        schemaXml = `<Field Type="Choice" Name="${name}" StaticName="${name}" DisplayName="${displayName}" Required="${required}" Format="Dropdown"><CHOICES>${choices}</CHOICES></Field>`;
        break;
      }
      case 'DateTime':
        schemaXml = `<Field Type="DateTime" Name="${name}" StaticName="${name}" DisplayName="${displayName}" Required="${required}" Format="DateOnly"/>`;
        break;
      case 'Number':
        schemaXml = `<Field Type="Number" Name="${name}" StaticName="${name}" DisplayName="${displayName}" Required="${required}"/>`;
        break;
      case 'Boolean':
        schemaXml = `<Field Type="Boolean" Name="${name}" StaticName="${name}" DisplayName="${displayName}"><Default>0</Default></Field>`;
        break;
      case 'User':
        schemaXml = `<Field Type="User" Name="${name}" StaticName="${name}" DisplayName="${displayName}" Required="${required}" UserSelectionMode="0"/>`;
        break;
      case 'URL':
        schemaXml = `<Field Type="URL" Name="${name}" StaticName="${name}" DisplayName="${displayName}" Required="${required}" Format="Hyperlink"/>`;
        break;
      default:
        schemaXml = `<Field Type="Text" Name="${name}" StaticName="${name}" DisplayName="${displayName}" MaxLength="255" Required="${required}"/>`;
    }

    await this.sp.web.lists
      .getByTitle(listName)
      .fields
      .createFieldAsXml(schemaXml);
  }

  private escapeXml(value: string): string {
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
