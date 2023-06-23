import { Environment, Log } from '@microsoft/sp-core-library';
import {
  ApplicationCustomizerContext, BaseApplicationCustomizer, PlaceholderContent, PlaceholderName,
} from '@microsoft/sp-application-base';
import { IReadonlyTheme, ISemanticColors } from '@microsoft/sp-component-base';

//import * as strings from 'AnnouncementsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'AnnouncementsApplicationCustomizer';

// Reference the solution
import "../../../../dist/announcements.js";
declare const Announcements: {
  render: (props: {
    el: HTMLElement;
    context?: ApplicationCustomizerContext;
    envType?: number;
    isEdit?: boolean;
  }) => void;
  updateBanner: (isEdit: boolean) => void;
  updateTheme: (currentTheme: Partial<ISemanticColors>) => void;
};

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAnnouncementsApplicationCustomizerProperties { }

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AnnouncementsApplicationCustomizer
  extends BaseApplicationCustomizer<IAnnouncementsApplicationCustomizerProperties> {

  private _banner: PlaceholderContent = null;

  public onInit(): Promise<void> {
    // Log
    Log.info(LOG_SOURCE, `Initializing the Announcements`);

    // Render the banner
    this.render();

    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    Announcements.updateTheme(currentTheme.semanticColors);
  }

  // Renders the banner
  private render(): void {
    // See if the banner has been created
    if (this._banner === null) {
      // Log
      Log.info(LOG_SOURCE, `Creating the top placeholder`);

      // Create the banner
      this._banner = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);
      this._banner.domElement.id = "announcements";
      this._banner.domElement.classList.add("bs");

      // Log
      Log.info(LOG_SOURCE, `Creating the banner`);

      // Render the banner
      Announcements.render({
        el: this._banner.domElement,
        context: this.context,
        envType: Environment.type
      });
    } else {
      // Log
      Log.info(LOG_SOURCE, `Banner already rendered`);
    }
  }
}
