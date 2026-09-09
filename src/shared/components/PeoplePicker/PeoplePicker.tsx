import * as React from 'react';
import { useMemo } from 'react';
import { NormalPeoplePicker, IPersonaProps, Label, ValidationState } from '@fluentui/react';
import { IUserValue } from '../../types/RhinoConfig';
import { PeopleService, IPersonSuggestion } from '../../services/PeopleService';
import { getAvatarColor, getInitials } from '../cells/format';

export interface IPeoplePickerProps {
  label?: string;
  required?: boolean;
  disabled?: boolean;
  multiple?: boolean;
  value: IUserValue[];
  onChange: (users: IUserValue[]) => void;
  peopleService?: PeopleService;
  errorMessage?: string;
  placeholder?: string;
}

interface IPersonaWithMeta extends IPersonaProps {
  suggestion?: IPersonSuggestion;
  user?: IUserValue;
}

function toPersona(u: IUserValue): IPersonaWithMeta {
  return {
    key: `u-${u.Id}`,
    text: u.Title || '',
    secondaryText: u.EMail || '',
    imageInitials: getInitials(u.Title || ''),
    initialsColor: undefined,
    user: u
  };
}

/** Fluent people picker bound to SharePoint's ClientPeoplePicker search. */
export const PeoplePicker: React.FC<IPeoplePickerProps> = ({
  label, required, disabled, multiple, value, onChange, peopleService, errorMessage, placeholder
}) => {
  const selected = useMemo(() => (value || []).map(toPersona), [value]);

  const resolve = async (filter: string, current?: IPersonaProps[]): Promise<IPersonaProps[]> => {
    if (!peopleService) return [];
    try {
      const people = await peopleService.search(filter);
      const taken: Record<string, boolean> = {};
      (current || []).forEach(p => { taken[String(p.secondaryText || p.text).toLowerCase()] = true; });
      return people
        .filter(p => !taken[(p.secondaryText || p.text).toLowerCase()])
        .map<IPersonaWithMeta>(p => ({
          key: p.key,
          text: p.text,
          secondaryText: p.secondaryText,
          imageInitials: getInitials(p.text),
          suggestion: p
        }));
    } catch {
      return [];
    }
  };

  const handleChange = async (items?: IPersonaProps[]): Promise<void> => {
    const list = items || [];
    const users: IUserValue[] = [];
    for (const p of list as IPersonaWithMeta[]) {
      if (p.user) { users.push(p.user); continue; }
      if (p.suggestion && peopleService) {
        try {
          users.push(await peopleService.ensure(p.suggestion));
        } catch {
          /* skip unresolved */
        }
      }
    }
    onChange(multiple ? users : users.slice(-1));
  };

  return (
    <div>
      {label && <Label required={required} disabled={disabled}>{label}</Label>}
      <NormalPeoplePicker
        onResolveSuggestions={resolve}
        onChange={items => { void handleChange(items); }}
        selectedItems={selected}
        itemLimit={multiple ? undefined : 1}
        disabled={disabled}
        inputProps={{ placeholder: placeholder || (multiple ? 'Search people…' : 'Search a person…') }}
        pickerSuggestionsProps={{ suggestionsHeaderText: 'People', noResultsFoundText: 'No matches', loadingText: 'Searching…' }}
        resolveDelay={250}
        onValidateInput={(input) => input && input.length > 1 ? ValidationState.valid : ValidationState.invalid}
        onRenderSuggestionsItem={(p) => (
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, padding: '6px 10px' }}>
            <span style={{ width: 28, height: 28, borderRadius: '50%', background: getAvatarColor(p.text || ''), color: '#fff', fontSize: 11, fontWeight: 600, display: 'inline-flex', alignItems: 'center', justifyContent: 'center' }}>{p.imageInitials}</span>
            <span style={{ display: 'flex', flexDirection: 'column' }}>
              <span style={{ fontSize: 13, fontWeight: 500 }}>{p.text}</span>
              <span style={{ fontSize: 11, color: '#616161' }}>{p.secondaryText}</span>
            </span>
          </div>
        )}
        styles={errorMessage ? { root: { selectors: { '.ms-BasePicker-text': { borderColor: '#a4262c' } } } } : undefined}
      />
      {errorMessage && <div style={{ color: '#a4262c', fontSize: 12, marginTop: 5 }}>{errorMessage}</div>}
    </div>
  );
};

export default PeoplePicker;
