import { Component } from "react";
import * as React from "react";
import styles from "./HelloWorld.module.scss";
import {
  Icon,
  IStackItemStyles,
  IStackStyles,
  IStackTokens,
  Stack,
} from "@fluentui/react";
const stackStyles: IStackStyles = {
  root: {
    background: "white",
    color: "black",
    // border: `1px solid #00000040`,
  },
};
const stackStyles2: IStackStyles = {
  root: {
    justifyContent: "space-between",
  },
};
const stackItemStyles: IStackItemStyles = {
  root: {
    fontFamily: "Segoe UI",
    fontSize: 16,
    color: `rgba(17, 28, 36, 1)`,
  },
};
const itemAlignmentsStackTokens: IStackTokens = {
  childrenGap: 20,
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
export default class Team extends Component {
  render() {
    return (
      <Stack
        enableScopedSelectors
        styles={stackStyles}
        tokens={itemAlignmentsStackTokens}
      >
        <div className={styles.teamsBanner}></div>
        <Stack
          horizontal={true}
          wrap={true}
          styles={stackStyles2}
          tokens={{ childrenGap: 20 }}
        >
          {users.map((item: any) => {
            return (
              <Stack>
               
                <Stack
                  className={styles.testTeamPage}
                  tokens={{ childrenGap: 14 }}
                >
                  <Stack tokens={{ childrenGap: 20 }} horizontal={true}>
                    <Stack>
                      <div className={styles.userImage} />
                    </Stack>
                    <Stack tokens={{ childrenGap: 8 }} styles={stackItemStyles}>
                      {/* <img src="C:\Users\Tejas\Documents\Test1\src\webparts\helloWorld\components\image.png" /> */}
                      <span className={styles.name}>{item.name}</span>
                      <span>{item.designation}</span>
                      <Stack
                        verticalAlign="center"
                        horizontal={true}
                        tokens={{ childrenGap: 4 }}
                      >
                        <Icon
                          aria-label="Compass"
                          iconName="Mail"
                          className={styles.iconStyle}
                        />
                        <span>{item.email}</span>
                      </Stack>
                      <Stack
                        horizontal={true}
                        tokens={{ childrenGap: 4 }}
                        verticalAlign="center"
                      >
                        <Icon
                          aria-label="Compass"
                          iconName="IncomingCall"
                          className={styles.iconStyle}
                          style={{
                            backgroundColor: "rgba(253, 183, 26, 1)",
                          }}
                        />
                        <span>{item.contactNo}</span>
                      </Stack>
                    </Stack>
                  </Stack>
                  <Stack.Item
                    style={{
                      fontSize: "18px",
                      fontFamily: "Segoe UI",
                      fontWeight: 400,
                      color: "rgba(17, 28, 36, 1)",
                      lineHeight: "26px",
                    }}
                  >
                    Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed
                    diam nonumy eirmod tempor invidunt ut labore et dolore magna
                    aliquyam erat, sed diam voluptua. At vero eos et accusam et
                    justo duo dolores et ea rebum.
                  </Stack.Item>
                </Stack>
                <div className={styles.testTeamPagebac}></div>
              </Stack>
            );
          })}
        </Stack>
      </Stack>
    );
  }
}
