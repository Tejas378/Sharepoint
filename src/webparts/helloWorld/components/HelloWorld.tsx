import * as React from "react";
import styles from "./HelloWorld.module.scss";
import type { IHelloWorldProps, IHelloWorlStates } from "./IHelloWorldProps";
import {
  Icon,
  IStackItemStyles,
  IStackStyles,
  IStackTokens,
  Separator,
  Stack,
} from "@fluentui/react";
import Team from "./Team";
const stackStyles: IStackStyles = {
  root: {
    background: "white",
    color: "black",
    // border: `1px solid #00000040`,
  },
};
const stackItemStyles: IStackItemStyles = {
  root: {
    // padding: 25,
    // backgroundColor: "pink",
    fontFamily: "Segoe UI",
    fontSize: 16,
    lineHeight: 21,
    color: `rgba(17, 28, 36, 1)`,
  },
};
const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 12,
  // padding: 10,
};
const users: any = [
  {
    name: "Tejas Sharma",
    email: "tejas.sharma@example.com",
    designation: "Software Engineer",
    contactNo: "+1-234-567-8901",
  },
  {
    name: "Maheshwari Patil",
    email: "maheshwari.patil@example.com",
    designation: "Frontend Developer",
    contactNo: "+1-345-678-9012",
  },
  {
    name: "Avdhut Kulkarni",
    email: "avdhut.kulkarni@example.com",
    designation: "React Developer",
    contactNo: "+1-456-789-0123",
  },
  {
    name: "Priya Mehta",
    email: "priya.mehta@example.com",
    designation: "Data Analyst",
    contactNo: "+1-567-890-1234",
  },
  {
    name: "Rahul Verma",
    email: "rahul.verma@example.com",
    designation: "Backend Developer",
    contactNo: "+1-678-901-2345",
  },
  {
    name: "Sneha Iyer",
    email: "sneha.iyer@example.com",
    designation: "Product Manager",
    contactNo: "+1-789-012-3456",
  },
  {
    name: "Amit Kumar",
    email: "amit.kumar@example.com",
    designation: "Cloud Engineer",
    contactNo: "+1-890-123-4567",
  },
  {
    name: "Nisha Gupta",
    email: "nisha.gupta@example.com",
    designation: "UI/UX Designer",
    contactNo: "+1-901-234-5678",
  },
  {
    name: "Vikram Singh",
    email: "vikram.singh@example.com",
    designation: "DevOps Engineer",
    contactNo: "+1-012-345-6789",
  },
  {
    name: "Kiran Joshi",
    email: "kiran.joshi@example.com",
    designation: "Business Analyst",
    contactNo: "+1-123-456-7890",
  },
];
export default class HelloWorld extends React.Component<
  IHelloWorldProps,
  IHelloWorlStates
> {
  constructor(props: IHelloWorldProps) {
    super(props);
    this.state = {  isDisplayTeams: false,}
  }
  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <section className={`${styles.helloWorld}`}>
        {this.state.isDisplayTeams!=null && this.state.isDisplayTeams!=false ? (
          <Team />
        ) : (
          <div className={styles.test}>
            <Stack
              enableScopedSelectors
              styles={stackStyles}
              tokens={itemAlignmentsStackTokens}
            >
              <Stack
                horizontal={true}
                tokens={{ childrenGap: 10 }}
                style={{ justifyContent: "space-between" }}
              >
                <Stack
                  horizontal={true}
                  verticalAlign="center"
                  tokens={{ childrenGap: 8 }}
                >
                  <Icon
                    iconName="People"
                    aria-label="Compass"
                    className={styles.peopleIconStayle}
                  />
                  <Stack.Item className={styles.EATeamHeaser}>
                    EA Team
                  </Stack.Item>
                </Stack>
                <Stack.Item
                  style={{ color: "rgba(90, 42, 130, 1)" }}
                  onClick={() => {
                    this.setState({ isDisplayTeams: true });
                  }}
                >
                  See all
                </Stack.Item>
              </Stack>
              <Separator />
              {users.map((item: any) => {
                return (
                  <>
                    <Stack horizontal={true} tokens={{ childrenGap: 20 }}>
                      <Stack>
                        <div className={styles.userImage} />
                      </Stack>
                      <Stack
                        tokens={{ childrenGap: 4 }}
                        styles={stackItemStyles}
                      >
                        {/* <img src="C:\Users\Tejas\Documents\Test1\src\webparts\helloWorld\components\image.png" /> */}
                        <span className={styles.name}>{item.name}</span>
                        <span>{item.designation}</span>
                        <Stack
                          verticalAlign="center"
                          horizontal={true}
                          tokens={{ childrenGap: 3 }}
                        >
                          <Icon
                            aria-label="Compass"
                            iconName="Mail"
                            style={{
                              color: "white",
                              backgroundColor: "rgba(0, 165, 171, 1)",
                              padding: "2px 4px",
                              borderRadius: 7,   fontSize:'12px'
                            }}
                          />
                          <span>{item.email}</span>
                        </Stack>
                        <Stack
                          horizontal={true}
                          tokens={{ childrenGap: 3 }}
                          verticalAlign="center"
                        >
                          <Icon
                            aria-label="Compass"
                            iconName="IncomingCall"
                            style={{
                              color: "white",
                              backgroundColor: "rgba(253, 183, 26, 1)",
                              padding: "2px 4px",
                              borderRadius: 7,
                              fontSize:'12px'
                            }}
                          />
                          <span>{item.designation}</span>
                        </Stack>
                      </Stack>
                    </Stack>
                    <Separator />
                  </>
                );
              })}
            </Stack>
          </div>
        )}
      </section>
    );
  }
}
