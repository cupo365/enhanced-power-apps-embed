import * as React from "react";
import styles from "./EnhancedPowerAppsEmbed.module.scss";
import { IEnhancedPowerAppsEmbedProps } from "./IEnhancedPowerAppsEmbedProps";
import * as strings from "EnhancedPowerAppsEmbedWebPartStrings";

/**
 * We use the placeholder to tell people they haven't configured the web part yet
 * */
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

/**
 * Used to provide support for theme variants
 */
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { IEnhancedPowerAppsEmbedState } from "./IEnhancedPowerAppsEmbedState";
import { IPowerAppsEmbedContext } from "../../../../models/IPowerAppsEmbedContext";

export default class EnhancedPowerApps extends React.Component<
  IEnhancedPowerAppsEmbedProps,
  IEnhancedPowerAppsEmbedState
> {
  public render(): React.ReactElement<IEnhancedPowerAppsEmbedProps> {
    const {
      dynamicProp,
      themeVariant,
      themeValues,
      appWebLink,
      useDynamicProp,
      dynamicPropName,
      locale,
      border,
      height,
      context,
    } = this.props;

    // The only thing we need for this web part to be configured is an app link or app id
    const needConfiguration: boolean = !appWebLink;

    const { semanticColors }: IReadonlyTheme = themeVariant;

    // If we passed a dynamic property, add it as a query string parameter
    const dynamicPropValue: string =
      useDynamicProp && dynamicProp !== undefined
        ? `&${encodeURIComponent(dynamicPropName)}=${encodeURIComponent(
            dynamicProp
          )}`
        : "";

    // We can take an app id or a full link. We'll assume (for now) that people are passing a valid app URL
    // would LOVE to find an API to retrieve list of valid apps
    const appUrl: string =
      appWebLink && appWebLink.indexOf("https://") != 0
        ? `https://apps.powerapps.com/play/${appWebLink}`
        : appWebLink;

    // Build the portion of the URL where we're passing theme colors
    let themeParams: string = "";

    if (themeValues && themeValues.length > 0) {
      themeValues.forEach((themeValue: string) => {
        try {
          const themeColor: string = semanticColors[themeValue];
          themeParams =
            themeParams + `&${themeValue}=${encodeURIComponent(themeColor)}`;
        } catch (e) {
          console.log(e);
        }
      });
    }

    // Some properties are commented out because the Param() function in Power Apps has a limit on the length of the URL (max 1000).
    // This means the embed context object cannot be too long. Current length frame URL is circa 800.
    const powerAppsEmbedContext: IPowerAppsEmbedContext = {
      aadInfo: {
        tenantId: context?.pageContext?.aadInfo?.tenantId?._guid
      },
      cultureInfo: {
        languageId: context?.pageContext?.web?.language,
        languageName: context?.pageContext?.web?.languageName
        //isRightToLeft: context?.pageContext?.cultureInfo?.isRightToLeft,
        // timeZone: {
        //   id: context?.pageContext?.legacyPageContext?.webTimeZoneData?.Id,
        //   name: context?.pageContext?.legacyPageContext?.webTimeZoneData?.Description,
        //   firstDayOfWeek: context?.pageContext?.legacyPageContext?.webFirstDayOfWeek
        // }
      },
      site: {
        title: context?.pageContext?.web?.title,
        portalUrl: context?.pageContext?.legacyPageContext?.portalUrl,
        absoluteUrl: context?.pageContext?.site?.absoluteUrl,
        serverRelativeUrl: context?.pageContext?.site?.serverRelativeUrl,
        serverRequestPath: context?.pageContext?.site?.serverRequestPath,
        group: {
          id: context?.pageContext?.legacyPageContext?.groupId,
          type: context?.pageContext?.legacyPageContext?.groupType
          //color: context?.pageContext?.legacyPageContext?.groupColor
        },
        isTeamsConnectedSite: context?.pageContext?.legacyPageContext?.isTeamsConnectedSite,
        isTeamsChannelSite: context?.pageContext?.legacyPageContext?.isTeamsChannelSite,
        isArchived: context?.pageContext?.legacyPageContext?.isArchived,
        isHubSite: context?.pageContext?.legacyPageContext?.isHubSite
        // sensitivityLabel: {
        //   id: context?.pageContext?.legacyPageContext?.sensitivityLabel,
        //   name: context?.pageContext?.legacyPageContext?.sensitivityLabelName
        // },
        // color: context?.pageContext?.legacyPageContext?.siteColor
      },
      user: {
        id: context?.pageContext?.aadInfo?.userId?._guid,
        displayName: context?.pageContext?.user?.displayName,
        email: context?.pageContext?.user?.email,
        isAnonymousGuestUser: context?.pageContext?.user?.isAnonymousGuestUser,
        isExternalGuestUser: context?.pageContext?.user?.isExternalGuestUser,
        isSiteAdmin: context?.pageContext?.legacyPageContext?.isSiteAdmin,
        isSiteOwner: context?.pageContext?.legacyPageContext?.isSiteOwner
        //userFirstDayOfWeek: context?.pageContext?.legacyPageContext?.userFirstDayOfWeek
      }
    };

    // Build the frame url
    const frameUrl: string = `${appUrl}?source=SPClient-EnhancedPowerAppsEmbed&amp;` +
      `locale=${locale}&amp;enableOnBehalfOf=true&amp;authMode=onbehalfof&amp;hideNavBar=true&amp;` +
      `${dynamicPropValue}${themeParams}&locale=${locale}` +
      `&tenantId=${powerAppsEmbedContext.aadInfo.tenantId}` +
      `&languageId=${powerAppsEmbedContext.cultureInfo.languageId}` +
      `&languageName=${powerAppsEmbedContext.cultureInfo.languageName}` +
      `&siteTitle=${powerAppsEmbedContext.site.title}` +
      `&portalUrl=${powerAppsEmbedContext.site.portalUrl}` +
      `&absoluteUrl=${powerAppsEmbedContext.site.absoluteUrl}` +
      `&serverRelativeUrl=${powerAppsEmbedContext.site.serverRelativeUrl}` +
      `&serverRequestPath=${powerAppsEmbedContext.site.serverRequestPath}` +
      `&groupId=${powerAppsEmbedContext.site.group.id}` +
      `&groupType=${powerAppsEmbedContext.site.group.type}` +
      `&isTeamsConnectedSite=${powerAppsEmbedContext.site.isTeamsConnectedSite}` +
      `&isTeamsChannelSite=${powerAppsEmbedContext.site.isTeamsChannelSite}` +
      `&isArchived=${powerAppsEmbedContext.site.isArchived}` +
      `&isHubSite=${powerAppsEmbedContext.site.isHubSite}` +
      `&userId=${powerAppsEmbedContext.user.id}` +
      `&displayName=${powerAppsEmbedContext.user.displayName}` +
      `&email=${powerAppsEmbedContext.user.email}` +
      `&isAnonymousGuestUser=${powerAppsEmbedContext.user.isAnonymousGuestUser}` +
      `&isExternalGuestUser=${powerAppsEmbedContext.user.isExternalGuestUser}` +
      `&isSiteAdmin=${powerAppsEmbedContext.user.isSiteAdmin}` +
      `&isSiteOwner=${powerAppsEmbedContext.user.isSiteOwner}`;
    console.log(frameUrl);

    return (
      <div
        className={styles.enhancedPowerApps}
        style={
          needConfiguration ? { height: `315px` } : { height: `${height}px` }
        }
      >
        {needConfiguration && (
          <Placeholder
            iconName="PowerApps"
            iconText={strings.PlaceholderIconText}
            description={strings.PlaceholderDescription}
            buttonLabel={strings.PlaceholderButtonLabel}
            onConfigure={this.props.onConfigure}
          />
        )}
        {!needConfiguration && (
          <>
            {this.props.appWebLink && (
              <iframe
                src={frameUrl}
                scrolling="no"
                allow="geolocation *; microphone *; camera *; fullscreen *;"
                sandbox="allow-popups allow-popups-to-escape-sandbox allow-same-origin allow-scripts allow-forms allow-orientation-lock allow-downloads"
                width={this.props.width}
                height={height}
                frameBorder={border ? "1" : "0"}
              ></iframe>
            )}
          </>
        )}
      </div>
    );
  }
}
