import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MyAssignmentsWebPartStrings';
import MyAssignments from './components/MyAssignments';
import { IMyAssignmentsWebPartProps,IMyAssignmentsProps } from './components/IMyAssignmentsProps';
import {
  IReadonlyTheme,ThemeProvider,ThemeChangedEventArgs
} from '@microsoft/sp-component-base';



export default class MyAssignmentsWebPart extends BaseClientSideWebPart<IMyAssignmentsWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  protected onInit(): Promise<void> {

    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);


    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IMyAssignmentsProps> = React.createElement(
      MyAssignments,
      {
        context: this.context,
        sphttpContext:this.context.spHttpClient,
        webPartProps:this.properties,
        themeVariant: this._themeVariant
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }



  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

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
                PropertyPaneSlider('pagingValue', {
                  label: "How many assignments would you like on each page?",
                  min:5,  
                  max:20,  
                  value:10,  
                  showValue:true,  
                  step:1 
                }),
                PropertyPaneToggle("subjectFilter", {
                  label: "Only show classes for this CDB automated subject site",
                  onText: "On",
                  offText: "Off",
                }),
                PropertyPaneToggle("hideOverDue", {
                  label: "Hide overdue assignments",
                  onText: "On",
                  offText: "Off",
                }),
                PropertyPaneToggle("showArchivedTeams", {
                  label: "Show assignments from archived class teams",
                  onText: "On",
                  offText: "Off",
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
