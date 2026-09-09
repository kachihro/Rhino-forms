import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/profiles';
import '@pnp/sp/site-users/web';
import { IUserValue } from '../types/RhinoConfig';

export interface IPersonSuggestion {
  key: string;        // login name / claims key
  text: string;       // display name
  secondaryText: string; // email
  id?: number;        // site user id once ensured
}

export class PeopleService {
  private cache: Record<string, IPersonSuggestion[]> = {};

  constructor(private sp: SPFI) {}

  public async search(query: string, max: number = 8): Promise<IPersonSuggestion[]> {
    const q = (query || '').trim();
    if (q.length < 2) return [];
    if (this.cache[q]) return this.cache[q];
    const results = await this.sp.profiles.clientPeoplePickerSearchUser({
      QueryString: q,
      MaximumEntitySuggestions: max,
      AllowEmailAddresses: true,
      AllowOnlyEmailAddresses: false,
      PrincipalSource: 15,
      PrincipalType: 1,
      SharePointGroupID: 0
    } as any);
    const people = results.map(r => ({
      key: r.Key,
      text: r.DisplayText,
      secondaryText: r.EntityData?.Email || r.Description || '',
      id: r.EntityData?.SPUserID ? Number(r.EntityData.SPUserID) : undefined
    }));
    this.cache[q] = people;
    return people;
  }

  /** Make sure the person exists in the site user list and return their Id. */
  public async ensure(person: IPersonSuggestion): Promise<IUserValue> {
    const result = await this.sp.web.ensureUser(person.key);
    const data: any = (result as any).data || result;
    return { Id: data.Id, Title: data.Title || person.text, EMail: data.Email || person.secondaryText };
  }
}
