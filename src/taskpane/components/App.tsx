import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: []
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Prime WebServer and  integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
  }

  click = async () => {
    function reqListener() {
      console.log(this.responseText);
    }
    try {
      await Excel.run(async context => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();
        const sheet = context.workbook.worksheets.getItem("Blad1");
        const range2 = sheet.getRange("A1:D5");
        range2.load("Values");

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "yellow";

        range2.select();
        await context.sync();
        console.log(`The range address was ${range.address}.`);
        console.log(`the other range2 was: ${JSON.stringify(range2.values)}`);
        const oReq = new XMLHttpRequest();
        oReq.addEventListener("load", reqListener);
        oReq.open("POST", "http://localhost:4000/api/projects", true);
        oReq.setRequestHeader("Content-Type", "application/json;charset=UTF-8");
        oReq.onreadystatechange = function() {
          console.log("hoho");
          if (oReq.readyState == 4 && oReq.status == 200) {
            console.log(oReq.responseText);
          }
        };
        oReq.send(JSON.stringify(range2.values));
        //console.log(JSON.stringify(range.text, null, 4));
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo="assets/CH_logo.png"
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header
          logo="assets/CH_logo.png"
          title={this.props.title}
          message="Welcome"
        />
        <HeroList
          message="CHYRONHEGO parkour add-in!"
          items={this.state.listItems}
        >
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Button1
          </Button>
        </HeroList>
      </div>
    );
  }
}
