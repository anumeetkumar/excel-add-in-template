import * as React from "react";
import TextInsertion from "./TextInsertion";
import { Button, makeStyles } from "@fluentui/react-components";
import { insertText, sumSelectedRange } from "../taskpane";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  console.log("props", props);
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  // const listItems: HeroListItem[] = [
  //   {
  //     icon: <Ribbon24Regular />,
  //     primaryText: "Achieve more with Office integration",
  //   },
  //   {
  //     icon: <LockOpen24Regular />,
  //     primaryText: "Unlock features and functionality",
  //   },
  //   {
  //     icon: <DesignIdeas24Regular />,
  //     primaryText: "Create and visualize like a pro",
  //   },
  // ];

  return (
    <div className={styles.root}>
      {/* <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems} /> */}
      <TextInsertion insertText={insertText} />
      <div style={{ padding: "20px" }}>
        
        <Button appearance="primary" disabled={false} size="large" onClick={sumSelectedRange}>
          Sum
        </Button>
      </div>
    </div>
  );
};

export default App;
