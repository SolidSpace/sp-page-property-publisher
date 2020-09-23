import * as React from 'react';
import { DisplayMode } from '@microsoft/sp-core-library';
import * as strings from 'PagePropertyPublisherWebPartStrings';
import { stringIsNullOrEmpty } from '@pnp/pnpjs';
//import styles from './Placeholder.module.scss';

export interface IPlaceholderProps {
  pageProperties: any;
  displaymode: DisplayMode;
  cbRefreshProperties: any;
}

export interface IPlaceholderState {
  displayList: boolean;
}

export class Placeholder extends React.Component<IPlaceholderProps, IPlaceholderState> {

  constructor(props: IPlaceholderProps) {
    super(props);

    this.state = {
      displayList: false
    };
  }

  private _propertyRenderer(): Array<JSX.Element> {
    let result: Array<JSX.Element> = new Array();
    let cnt = 0;
    for (var key in this.props.pageProperties) {
      let value: any = this.props.pageProperties[key];
      value = (typeof value == 'object') ? JSON.stringify(value) : value;
      let rowCss = (cnt % 2) ? "item-even" : "item-odd";
      if (key != "CanvasContent1" && key != "LayoutWebpartsContent") {
        cnt++;
        result.push(<div className={"se-pageproperty-listitem " + rowCss}><div className="se-pageproperty-listitem-key">{key}</div><div className="se-pageproperty-listitem-value">{(value == null) ? "null" : value}</div></div>);
      }

    }
    return result;
  }

  private _setDisplayList() {
    this.setState({ ...this.state, displayList: !this.state.displayList });
  }

  public render(): React.ReactElement<IPlaceholderProps> {

    var cssClass = (this.props.displaymode == DisplayMode.Edit) ? "se-placeholder-wrapper se-placeholder-editmode" : "se-placeholder-readmode";
    return (
      <div className={cssClass}>
        <div className="se-pageproperty-wrapper">
          <div className="se-pageproperty-header">
            <div className="se-pageproperty-headertext">Page Property Publisher</div>
            <div className="se-pageproperty-headericons">
              <div className="se-pageproperty-icon" onClick={() => this.props.cbRefreshProperties(true)}><i className="ms-Icon ms-Icon--Refresh ms-fontSize-18" aria-hidden="true"></i></div>
              <div className={(!this.state.displayList) ? "se-pageproperty-icon show" : "se-pageproperty-icon hide"} onClick={this._setDisplayList.bind(this)}>
                <i className="ms-Icon ms-Icon--ChevronDown ms-fontSize-18" aria-hidden="true"></i>
              </div>
              <div className={(this.state.displayList) ? "se-pageproperty-icon show" : "se-pageproperty-icon hide"} onClick={this._setDisplayList.bind(this)}>
                <i className="ms-Icon ms-Icon--ChevronUp ms-fontSize-18" aria-hidden="true"></i>
              </div>
            </div>
          </div>
          <div className={(!this.state.displayList) ? "se-pageproperty-list show" : "se-pageproperty-list hide"}>
            <div className="se-pageproperty-about-header">
                {strings.AboutHeadline}
            </div>
            <div className="se-pageproperty-about-text">
              {strings.AboutText}
            </div>
            <div className="se-pageproperty-about-header">
            {strings.UsingHeadline}
            </div>
            <div className="se-pageproperty-about-text">
              {strings.UsingText}
            </div>
          </div>
          <div className={(this.state.displayList && this.props.pageProperties.Title == undefined) ? "se-pageproperty-list show" : "se-pageproperty-list hide"} onClick={() => this.props.cbRefreshProperties(true)}>
            <div className="se-pageproperties-erricon">
              <i className="ms-Icon ms-Icon--Robot" aria-hidden="true"></i>
            </div>
            <div  className="se-pageproperties-errtext">
              {strings.ErrorShowPropertiesText}
              </div>
          </div>
          <div className={(this.state.displayList) ? "se-pageproperty-list show" : "se-pageproperty-list hide"}>
            {(this.props.pageProperties.Title != undefined) ? this._propertyRenderer() : ""}
          </div>
        </div>
      </div>
    );
  }
}
