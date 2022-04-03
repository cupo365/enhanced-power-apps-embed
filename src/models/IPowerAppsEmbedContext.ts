export interface IPowerAppsEmbedContext {
  aadInfo?: IAADInfoContext;
  cultureInfo?: ICultureInfoContext;
  site?: ISiteContext;
  user?: IUserContext;
}

export interface IAADInfoContext {
  tenantId?: string;
}

export interface ICultureInfoContext {
  languageId: number;
  languageName?: string;
  isRightToLeft?: boolean;
  timeZone?: ITimeZoneContext;
}

export interface ITimeZoneContext {
  id?: number;
  name?: string;
  firstDayOfWeek?: FirstDayOfWeek;
}

export enum FirstDayOfWeek {
  Sunday = 0,
  Monday = 1,
  Tuesday = 2,
  Wednesday = 3,
  Thursday = 4,
  Friday = 5,
  Saturday = 6
}

export interface ISiteContext {
  title?: string;
  portalUrl?: string;
  absoluteUrl?: string;
  serverRelativeUrl?: string;
  serverRequestPath?: string;
  group?: IGroupContext;
  isTeamsConnectedSite?: boolean;
  isTeamsChannelSite?: boolean;
  isArchived?: boolean;
  isHubSite?: boolean;
  sensitivityLabel?: ISensitivityLabelContext;
  color?: string;
}

export interface IGroupContext {
  id?: string;
  type?: string;
  color?: string;
}

export interface ISensitivityLabelContext {
  id?: string;
  name?: string;
}

export interface IUserContext {
  id?: string;
  displayName?: string;
  email?: string;
  isAnonymousGuestUser?: boolean;
  isExternalGuestUser?: boolean;
  isSiteAdmin?: boolean;
  isSiteOwner?: boolean;
  userFirstDayOfWeek?: FirstDayOfWeek;
}
