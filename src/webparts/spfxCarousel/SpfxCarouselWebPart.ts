import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'SpfxCarouselWebPartStrings';
import SpfxCarousel from './components/SpfxCarousel';
import { ICarouselSlide, ISpfxCarouselProps } from './components/ISpfxCarouselProps';

interface IRenderedListDataResponse {
  Row?: IRenderedNewsItem[];
}

interface IIdResponse {
  Id: string;
}

type IRenderedNewsItem = Record<string, unknown> & {
  ID?: string;
  Id?: string;
  Title?: string;
  Body?: string;
};

interface IListContextInfo {
  // These IDs are needed to build SharePoint's thumbnail API URL for list item images.
  listId: string;
  siteId: string;
  webId: string;
}

export interface ISpfxCarouselWebPartProps {
  siteUrl: string;
  enableAutoplay: boolean;
  autoplayDelaySeconds: number;
  slideLimit: number;
}

export default class SpfxCarouselWebPart extends BaseClientSideWebPart<ISpfxCarouselWebPartProps> {
  // The News list is pinned to a known GUID so a future rename does not break the web part.
  private static readonly _newsListId: string = '7b68641e-c9b4-48c5-831b-04938fdcce43';
  // SharePoint's rendered payload can expose the same logical field under different internal names.
  // We keep those variants in one place so mapping stays easy to adjust.
  private static readonly _itemIdFieldIds: string[] = ['ID', 'Id'];
  private static readonly _titleFieldIds: string[] = ['Title', 'Title.','Headline', 'Headline_x0020_Suggestion'];
  private static readonly _bodyFieldIds: string[] = ['Body', 'Body.', 'Story_x0020_Details'];
  private static readonly _targetUrlFieldIds: string[] = [
    'Target URL',
    'Target_x0020_URL',
    'TargetURL',
    'TargetUrl',
    'Link'
  ];
  private static readonly _imageFieldIds: string[] = ['Image.', 'Image', 'Image0', 'Picture'];
  private static readonly _destinationFieldIds: string[] = [
    'News Destination',
    'News_x0020_Destination',
    'NewsDestination',
    'Destination'
  ];
  private _isDarkTheme: boolean = false;
  private _slides: ICarouselSlide[] = [];
  private _isLoading: boolean = false;
  private _errorMessage: string | undefined;

  public render(): void {
    // The web part owns SharePoint data loading; the React component is only responsible for presentation.
    const element: React.ReactElement<ISpfxCarouselProps> = React.createElement(
      SpfxCarousel,
      {
        slides: this._slides,
        enableAutoplay: this.properties.enableAutoplay === true,
        autoplayDelaySeconds: this.properties.autoplayDelaySeconds || 5,
        isLoading: this._isLoading,
        errorMessage: this._errorMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await this._loadSlides(this._normalizeSiteUrl(this.properties.siteUrl));
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath === 'siteUrl' && oldValue !== newValue) {
      this._loadSlides(this._normalizeSiteUrl(newValue as string)).catch(() => {
        // _loadSlides manages its own error state for rendering.
      });
    }

    if (propertyPath === 'slideLimit' && oldValue !== newValue) {
      this._loadSlides(this._normalizeSiteUrl(this.properties.siteUrl)).catch(() => {
        // _loadSlides manages its own error state for rendering.
      });
    }
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: 'Site URL',
                  description: 'The SharePoint site that contains the News list.'
                }),
                PropertyPaneToggle('enableAutoplay', {
                  label: 'Autoplay',
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneSlider('autoplayDelaySeconds', {
                  label: 'Autoplay delay (seconds)',
                  min: 2,
                  max: 10,
                  step: 1,
                  value: this.properties.autoplayDelaySeconds || 5,
                  showValue: true
                }),
                PropertyPaneSlider('slideLimit', {
                  label: 'Maximum slides',
                  min: 1,
                  max: 12,
                  step: 1,
                  value: this.properties.slideLimit || 5,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private async _loadSlides(siteUrl: string | undefined): Promise<void> {
    if (!siteUrl) {
      this._slides = [];
      this._errorMessage = 'Add the site URL for the SharePoint site that contains the News list.';
      this._isLoading = false;
      this.render();
      return;
    }

    this._isLoading = true;
    this._errorMessage = undefined;
    this.render();

    try {
      const listContext: IListContextInfo = await this._getListContextInfo(siteUrl);
      // RenderListDataAsStream gives us SharePoint's rendered field output, which is more useful here than the raw
      // list item endpoint because image and link fields often come back as preview HTML/button markup.
      const requestUrl: string = `${siteUrl}/_api/web/lists(guid'${listContext.listId}')/RenderListDataAsStream`;
      const requestBody: string = JSON.stringify({
        parameters: {
          RenderOptions: 2,
          // We only need a simple most-recent-first slice for the carousel.
          ViewXml:
            '<View><Query><OrderBy><FieldRef Name="Created" Ascending="FALSE" /></OrderBy></Query>' +
            `<RowLimit>${this.properties.slideLimit || 5}</RowLimit></View>`
        }
      });

      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        requestUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata'
          },
          body: requestBody
        }
      );

      if (!response.ok) {
        throw new Error(`SharePoint returned ${response.status} ${response.statusText}`);
      }

      const payload: IRenderedListDataResponse = await response.json();
      const items: IRenderedNewsItem[] = Array.isArray(payload.Row) ? payload.Row : [];

      this._slides = items
        // Only items explicitly tagged for the rotator should appear in this web part.
        .filter((item: IRenderedNewsItem) => this._isRotatorItem(item))
        .map((item: IRenderedNewsItem) => this._mapItemToSlide(item, siteUrl, listContext))
        .filter((slide: ICarouselSlide | undefined): slide is ICarouselSlide => slide !== undefined)
        .slice(0, this.properties.slideLimit || 5);

      this._errorMessage = undefined;
    } catch (error) {
      this._slides = [];
      this._errorMessage = error instanceof Error
        ? `Unable to load News list items. ${error.message}`
        : 'Unable to load News list items.';
    } finally {
      this._isLoading = false;
      this.render();
    }
  }

  private _mapItemToSlide(
    item: IRenderedNewsItem,
    siteUrl: string,
    listContext: IListContextInfo
  ): ICarouselSlide | undefined {
    // We read the rendered field values first because SharePoint sometimes exposes richer HTML/preview payloads
    // there than in the raw field output.
    const rawTargetField: unknown = this._getFirstDefinedField(item, SpfxCarouselWebPart._targetUrlFieldIds);
    const rawImageField: unknown = this._getFirstDefinedField(item, SpfxCarouselWebPart._imageFieldIds);
    const title: string = this._toFieldText(this._getFirstDefinedField(item, SpfxCarouselWebPart._titleFieldIds)) || 'Untitled news item';
    const summary: string = this._toPlainText(this._toFieldText(this._getFirstDefinedField(item, SpfxCarouselWebPart._bodyFieldIds)));
    const href: string | undefined = this._extractHref(rawTargetField, siteUrl);
    // Uploaded list images are not stored as plain URLs, so we rebuild the thumbnail URL when needed.
    const attachmentImageUrl: string | undefined = this._getImageAttachmentUrl(rawImageField, item, siteUrl, listContext);
    const imageSrc: string | undefined =
      this._extractImageSrc(rawImageField, siteUrl) ||
      attachmentImageUrl ||
      this._extractImageSrc(this._getFirstDefinedField(item, SpfxCarouselWebPart._bodyFieldIds), siteUrl);

    return {
      title,
      summary,
      href,
      imageSrc,
      imageAlt: title
    };
  }

  private _isRotatorItem(item: IRenderedNewsItem): boolean {
    const destination: string = this._toFieldText(
      this._getFirstDefinedField(item, SpfxCarouselWebPart._destinationFieldIds)
    ).toLowerCase();

    return destination.indexOf('rotator') > -1;
  }

  private _getFirstDefinedField(item: IRenderedNewsItem, fieldNames: string[]): unknown {
    // SharePoint field naming can differ between rendered output and internal field names, so we try each known alias.
    for (const fieldName of fieldNames) {
      if (Object.prototype.hasOwnProperty.call(item, fieldName) && item[fieldName] !== undefined && item[fieldName] !== null) {
        return item[fieldName];
      }
    }

    return undefined;
  }

  private _extractHref(value: unknown, siteUrl: string): string | undefined {
    if (!value) {
      return undefined;
    }

    if (typeof value === 'string') {
      // Hyperlink fields often come back as rendered anchor HTML instead of a bare URL.
      const hrefMatch: RegExpMatchArray | null = value.match(/href=["']([^"']+)["']/i);
      if (hrefMatch?.[1]) {
        return this._toAbsoluteUrl(hrefMatch[1], siteUrl);
      }

      const trimmedValue: string = value.trim();
      if (trimmedValue.indexOf('http') === 0 || trimmedValue.indexOf('/') === 0) {
        return this._toAbsoluteUrl(trimmedValue, siteUrl);
      }
    }

    if (value && typeof value === 'object') {
      const recordValue: Record<string, unknown> = value as Record<string, unknown>;
      const candidateHref: unknown = recordValue.Url || recordValue.url || recordValue.href;

      if (typeof candidateHref === 'string') {
        return this._toAbsoluteUrl(candidateHref, siteUrl);
      }
    }

    return undefined;
  }

  private _extractImageSrc(value: unknown, siteUrl: string): string | undefined {
    if (!value) {
      return undefined;
    }

    if (typeof value === 'string') {
      // SharePoint sometimes buries the real image URL inside preview JSON rather than a normal img tag.
      const imagePreviewMatch: RegExpMatchArray | null = value.match(/"imagePreview"\s*:\s*"([^"]+)"/i);
      if (imagePreviewMatch?.[1]) {
        return this._toAbsoluteUrl(this._decodeEscapedUrl(imagePreviewMatch[1]), siteUrl);
      }

      const srcMatch: RegExpMatchArray | null = value.match(/src=["']([^"']+)["']/i);
      if (srcMatch?.[1]) {
        return this._toAbsoluteUrl(srcMatch[1], siteUrl);
      }

      const trimmedValue: string = value.trim();
      if (trimmedValue.indexOf('http') === 0 || trimmedValue.indexOf('/') === 0) {
        return this._toAbsoluteUrl(trimmedValue, siteUrl);
      }

      try {
        const parsedValue: unknown = JSON.parse(trimmedValue);
        return this._extractImageSrc(parsedValue, siteUrl);
      } catch {
        return undefined;
      }
    }

    if (Array.isArray(value)) {
      for (const entry of value) {
        const imageSrc: string | undefined = this._extractImageSrc(entry, siteUrl);
        if (imageSrc) {
          return imageSrc;
        }
      }

      return undefined;
    }

    if (typeof value === 'object') {
      const recordValue: Record<string, unknown> = value as Record<string, unknown>;
      // Prefer direct URL-like properties before recursively walking the rest of the payload.
      const directCandidate: unknown =
        recordValue.src ||
        recordValue.Url ||
        recordValue.url ||
        recordValue.serverRelativeUrl;

      if (typeof directCandidate === 'string') {
        return this._toAbsoluteUrl(directCandidate, siteUrl);
      }

      for (const key in recordValue) {
        if (!Object.prototype.hasOwnProperty.call(recordValue, key)) {
          continue;
        }

        const imageSrc: string | undefined = this._extractImageSrc(recordValue[key], siteUrl);
        if (imageSrc) {
          return imageSrc;
        }
      }
    }

    return undefined;
  }

  private _getImageAttachmentUrl(
    value: unknown,
    item: IRenderedNewsItem,
    siteUrl: string,
    listContext: IListContextInfo
  ): string | undefined {
    let parsedValue: { fileName?: string } | undefined;

    if (typeof value === 'string') {
      try {
        parsedValue = JSON.parse(value) as { fileName?: string };
      } catch {
        return undefined;
      }
    } else if (value && typeof value === 'object') {
      parsedValue = value as { fileName?: string };
    } else {
      return undefined;
    }

    const fileName: string | undefined = parsedValue.fileName;
    const rawItemId: unknown = this._getFirstDefinedField(item, SpfxCarouselWebPart._itemIdFieldIds);
    const itemId: string =
      typeof rawItemId === 'number'
        ? String(rawItemId)
        : this._toFieldText(rawItemId);

    if (!fileName || !itemId) {
      return undefined;
    }

    // SharePoint-hosted list images are served from the v2 thumbnail API, not from the file name alone.
    return `${siteUrl}/_api/v2.1/sites('${listContext.siteId},${listContext.webId}')/lists('${listContext.listId}')/items('${itemId}')/attachments('${encodeURIComponent(fileName)}')/thumbnails/0/c1600x900/content?prefer=noredirect,closestavailablesize`;
  }

  private async _getListContextInfo(siteUrl: string): Promise<IListContextInfo> {
    // These three IDs are enough to reconstruct image thumbnail URLs for uploaded list images.
    const [siteResponse, webResponse, listResponse] = await Promise.all([
      this.context.spHttpClient.get(
        `${siteUrl}/_api/site?$select=Id`,
        SPHttpClient.configurations.v1
      ),
      this.context.spHttpClient.get(
        `${siteUrl}/_api/web?$select=Id`,
        SPHttpClient.configurations.v1
      ),
      this.context.spHttpClient.get(
        `${siteUrl}/_api/web/lists(guid'${SpfxCarouselWebPart._newsListId}')?$select=Id`,
        SPHttpClient.configurations.v1
      )
    ]);

    if (!siteResponse.ok || !webResponse.ok || !listResponse.ok) {
      throw new Error('Unable to resolve site or list IDs for News attachments.');
    }

    const sitePayload: IIdResponse = await siteResponse.json();
    const webPayload: IIdResponse = await webResponse.json();
    const listPayload: IIdResponse = await listResponse.json();

    return {
      siteId: sitePayload.Id,
      webId: webPayload.Id,
      listId: listPayload.Id
    };
  }

  private _toAbsoluteUrl(value: string, siteUrl: string): string {
    // Rendered list payloads can return either absolute URLs or site-relative URLs.
    try {
      return new URL(value, siteUrl).toString();
    } catch {
      return value;
    }
  }

  private _decodeEscapedUrl(value: string): string {
    // SharePoint preview JSON escapes slashes in some responses; decode them before using the URL.
    return value
      .replace(/\\u002f/gi, '/')
      .replace(/\\\//g, '/')
      .replace(/&amp;/gi, '&');
  }

  private _toPlainText(value: string | undefined): string {
    if (!value) {
      return '';
    }

    // Body is rich text in the list; strip markup so the carousel shows a clean summary.
    return value
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private _toFieldText(value: unknown): string {
    // Rendered list output is inconsistent across field types, so this normalizes the common text containers.
    if (typeof value === 'string') {
      return value;
    }

    if (Array.isArray(value)) {
      return value.map((entry: unknown) => this._toFieldText(entry)).join(' ');
    }

    if (value && typeof value === 'object') {
      const candidateObject: Record<string, unknown> = value as Record<string, unknown>;
      return [
        candidateObject.Label,
        candidateObject.Value,
        candidateObject.Title,
        candidateObject.lookupValue
      ]
        .map((entry: unknown) => (typeof entry === 'string' ? entry : ''))
        .join(' ');
    }

    return '';
  }

  private _normalizeSiteUrl(value: string | undefined): string | undefined {
    const trimmedValue: string = value?.trim() || '';
    return trimmedValue ? trimmedValue.replace(/\/$/, '') : undefined;
  }
}
