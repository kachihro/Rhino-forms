import { IRhinoConfig, IGridItem, IFieldDefinition } from '../types/RhinoConfig';

const PEOPLE = [
  { Id: 11, Title: 'Amelia Chen', EMail: 'amelia.chen@example.com' },
  { Id: 12, Title: 'Noah Patel', EMail: 'noah.patel@example.com' },
  { Id: 13, Title: 'Olivia Nguyen', EMail: 'olivia.nguyen@example.com' },
  { Id: 14, Title: 'Liam O\'Brien', EMail: 'liam.obrien@example.com' },
  { Id: 15, Title: 'Isla Thompson', EMail: 'isla.thompson@example.com' }
];

const WORDS = ['Website refresh', 'Q3 budget review', 'Onboarding pack', 'Warehouse audit', 'Fleet renewal', 'Client portal', 'Safety training', 'Data migration', 'Brand guidelines', 'Office move'];
const LOREM = 'Follow-up required with the stakeholders before the next milestone. Notes captured during the review meeting.';

function pick<T>(arr: T[], i: number): T { return arr[i % arr.length]; }

function sampleValue(field: IFieldDefinition, i: number): unknown {
  switch (field.type) {
    case 'Text': return field.internalName === 'Title' ? pick(WORDS, i) : `${field.displayName} ${i + 1}`;
    case 'Note': return LOREM;
    case 'Choice': return field.choices && field.choices.length ? pick(field.choices, i) : undefined;
    case 'MultiChoice': return field.choices && field.choices.length ? field.choices.filter((_, idx) => (idx + i) % 2 === 0).slice(0, 2) : [];
    case 'DateTime': { const d = new Date(); d.setDate(d.getDate() + (i * 7) - 14); return d.toISOString(); }
    case 'Number': return field.numberFormat === 'percent' ? (i * 17) % 100 : field.numberFormat === 'integer' ? (i + 1) * 3 : ((i + 1) * 1250.5);
    case 'Boolean': return i % 2 === 0;
    case 'User': return field.allowMultiple ? [pick(PEOPLE, i), pick(PEOPLE, i + 2)] : pick(PEOPLE, i);
    case 'URL': return { Url: 'https://example.com', Description: 'example.com' };
    default: return '';
  }
}

/** Fake rows for the designer's live previews. */
export function sampleItems(config: IRhinoConfig, count: number = 6): IGridItem[] {
  const items: IGridItem[] = [];
  for (let i = 0; i < count; i++) {
    const item: IGridItem = { Id: i + 1 };
    config.fields.forEach(f => { if (!f.isSystem) item[f.internalName] = sampleValue(f, i); });
    const created = new Date(); created.setDate(created.getDate() - (count - i) * 3);
    item.Created = created.toISOString();
    item.Modified = new Date().toISOString();
    item.Author = pick(PEOPLE, i);
    item.Editor = pick(PEOPLE, i + 1);
    item.ID = i + 1;
    if (item.Title === undefined) item.Title = pick(WORDS, i);
    items.push(item);
  }
  return items;
}
