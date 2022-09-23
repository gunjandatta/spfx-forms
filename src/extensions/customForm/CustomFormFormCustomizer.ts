import { Environment, Log } from "@microsoft/sp-core-library";
import {
  BaseFormCustomizer
} from '@microsoft/sp-listview-extensibility';

import * as strings from 'CustomFormFormCustomizerStrings';
import styles from './CustomFormFormCustomizer.module.scss';

import { WebParts } from "gd-sprest-bs";

/**
 * If your form customizer uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomFormFormCustomizerProperties { }

const LOG_SOURCE: string = 'CustomFormFormCustomizer';

export default class CustomFormFormCustomizer
  extends BaseFormCustomizer<ICustomFormFormCustomizerProperties> {

  public onInit(): Promise<void> {
    // Add your custom initialization to this method. The framework will wait
    // for the returned promise to resolve before rendering the form.
    Log.info(LOG_SOURCE, 'Activated CustomFormFormCustomizer with properties:');
    Log.info(LOG_SOURCE, JSON.stringify(this.properties, undefined, 2));
    return Promise.resolve();
  }

  public render(): void {
    // See if we have already rendered the form
    if (this.domElement.querySelector("form")) { return; }

    // Render the custom form webpart
    WebParts.SPFxListFormWebPart({
      envType: Environment.type,
      spfx: this as any
    });
  }

  public onDispose(): void {
    // This method should be used to free any resources that were allocated during rendering.
    super.onDispose();
  }
}
