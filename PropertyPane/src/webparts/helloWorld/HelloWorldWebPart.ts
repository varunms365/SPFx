import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneLabel, //Label
  PropertyPaneTextField, //Textbox
  PropertyPaneLink, //Link
  PropertyPaneDropdown, //Dropdown
  PropertyPaneCheckbox, //Checkbox
  PropertyPaneChoiceGroup, //Choice
  PropertyPaneToggle, //Toggle
  PropertyPaneSlider, //Slider
  PropertyPaneButton, //Button
  PropertyPaneHorizontalRule, //HorizontalRule
  PropertyPaneButtonType 
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

export interface IHelloWorldWebPartProps {
  name: string;
  description: string;
  textlabel: string;
  url: string;
  dropdown: string;
  checkbox: string;
  choice: string;
  toggle: string;
  slider: string;
  button: string;
  horizontalrule: string;
}

/*protected get disableReactivePropertyChanges(): boolean {
  return true;
}*/

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>

              <p class="${ styles.description }">Name : ${escape(this.properties.name)}</p>
              <p class="${ styles.description }">Description : ${escape(this.properties.description)}</p>
              <p class="${ styles.description }">Textlabel : ${escape(this.properties.textlabel)}</p>
              <p class="${ styles.description }">url : ${escape(this.properties.url)}</p>

              <p class="${ styles.description }">horizontalrule : ${escape(this.properties.horizontalrule)}</p>

              <p class="${ styles.description }">dropdown : ${escape(this.properties.dropdown)}</p>
              <p class="${ styles.description }">checkbox : ${escape(this.properties.checkbox)}</p>
              <p class="${ styles.description }">choice : ${escape(this.properties.choice)}</p>

              <p class="${ styles.description }">horizontalrule : ${escape(this.properties.horizontalrule)}</p>

              <p class="${ styles.description }">slider : ${escape(this.properties.slider)}</p>
              <p class="${ styles.description }">toggle : ${escape(this.properties.toggle)}</p>

              <p class="${ styles.description }">button : ${escape(this.properties.button)}</p>
              <p class="${ styles.description }">horizontalrule : ${escape(this.properties.horizontalrule)}</p>

              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected TextBoxValidationMethod(value: string): string{
    if(value.length <= 4) {return "Name should be at least 5 charecters.";}
    else { return ''; }
  }
  protected ButtonClick(val: any): any{
    this.properties.name = "Button Click Succeeded";
    return "test";
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Property Pane Page 1 - Name, Description, Label and Link"
          },
          groups: [
            {
              groupName: "Property Pane Group 1 Page 1", 
              groupFields: [
                PropertyPaneTextField('name', {
                  label: "Name",
                  multiline:false,
                  resizable:false,
                  onGetErrorMessage:this.TextBoxValidationMethod,
                  errorMessage:"This is error message for Name",
                  deferredValidationTime:10000,
                  placeholder:"Please enter name",
                  "description": "Name property pane field"
                }),
                PropertyPaneTextField('description', {
                  label: "Description",
                  multiline: true,
                  resizable:true,
                  placeholder:"Please enter description",
                  "description": "Description property pane field"
                }),
                PropertyPaneLabel('label', {
                  text:"Please enter text",
                  required:true
                }),
                PropertyPaneTextField("textlabel",{}),
                PropertyPaneLink('url', {
                  href: 'http://VarunAtluri.com',
                  text: "Varun Atluri's blog",
                  target: '_blank',
                  popupWindowProps: {
                    height: 500,
                    width: 500,
                    positionWindowPosition: 2,
                    title: "Varun Atluri's blog"
                  }
              })
              ]
            }
          ]
        },
        {
          header: {
            description: "Property Pane Page 2 - Dropdown, Checkbox and choice"
          },
          groups: [
            {
              groupName: "Property Pane Group 1 Page 2", 
              groupFields: [
                PropertyPaneDropdown('dropdown', {
                  label: "Drop Down",
                  options:[
                    {key: 'Drop Down 1', text:'Drop Down 1'},
                    {key: 'Drop Down 2', text:'Drop Down 2'},
                    {key: 'Drop Down 3', text:'Drop Down 3'}
                  ],
                  selectedKey: 'Drop Down 2'
                }),
                PropertyPaneCheckbox('checkbox', {
                  text: "Yes/No",
                  checked:true,
                  disabled:false
                }),
              ]
            },
            {
              groupName: "Property Pane Group 2 Page 2", 
              groupFields: [
                PropertyPaneChoiceGroup('choice', {
                  label: 'Choices',
                  options: [
                    { key: 'Choice 1', text: 'Choice 1' },
                    { key: 'Choice 2', text: 'Choice 2', checked: true },
                    { key: 'Choice 3', text: 'Choice 3' }
                  ],
                  
                })
              ]
            }
          ]
        },
        {
          header: {
            description: "Property Pane Page 3 - slider, toggle, button and Horizontal rule" 
          },
          groups: [
            {
              groupName: "Property Pane Group 1 Page 3", 
              groupFields: [
                PropertyPaneSlider('slider', {
                  label: "Slider",
                  min:1,
                  max:5
                }),
                PropertyPaneToggle('toggle', {
                  label: "Slider",
                })
              ]
            },
            {
              groupName: "Property Pane Group 2 Page 3", 
              groupFields: [
                PropertyPaneButton('Button', {
                  text: "Normal button",
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this.ButtonClick.bind(this)
                 }),
                 PropertyPaneHorizontalRule(),
              ]
            }
          ]
        }
      ]
    };
  }
}
