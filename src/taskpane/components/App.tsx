import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global console, Office, require */

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
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  click = async () => {
    /**
     * Insert your PowerPoint code here
     */
    Office.context.document.setSelectedDataAsync(
      "Hello World!",
      {
        coercionType: Office.CoercionType.Text,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  };

  insertBox = async () => {
    // This sample creates a rectangle positioned 100 points from the top and left sides
    // of the slide and is 150x150 points. The shape is put on the first slide.
    await PowerPoint.run(async (context) => {
      await context.sync();
      // @ts-ignore
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.diamond);
      rectangle.left = 100;
      rectangle.top = 100;
      rectangle.height = 150;
      rectangle.width = 150;
      rectangle.fill.color.rgb = "FF0000";
      rectangle.name = "Square";
      await context.sync();
    });
  };

  deleteShapes = async () => {
    await PowerPoint.run(async (context) => {
      // Delete all shapes from the first slide.
      const sheet = context.presentation.slides.getItemAt(0);
      // @ts-ignore
      const shapes = sheet.shapes;

      // Load all the shapes in the collection without loading their properties.
      shapes.load("items/$none");
      await context.sync();

      shapes.items.forEach(function (shape) {
        shape.delete();
      });
      await context.sync();
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.click}>
            Run
          </DefaultButton>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronLeft" }} onClick={this.insertBox}>
            Add a box
          </DefaultButton>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronLeft" }} onClick={this.deleteShapes}>
            Delete all shapes
          </DefaultButton>
      </div>
    );
  }
}
