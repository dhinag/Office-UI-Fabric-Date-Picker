import * as React from "react";
import * as Sugar from "sugar";

import {
  DatePicker,
  IDatePickerStrings
} from "office-ui-fabric-react/lib/DatePicker";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";

initializeIcons(undefined, { disableWarnings: true });

const DayPickerStrings: IDatePickerStrings = {
  months: [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
  ],

  shortMonths: [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec"
  ],

  days: [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday"
  ],

  shortDays: ["S", "M", "T", "W", "T", "F", "S"],

  goToToday: "Go to today",
  prevMonthAriaLabel: "Go to previous month",
  nextMonthAriaLabel: "Go to next month",
  prevYearAriaLabel: "Go to previous year",
  nextYearAriaLabel: "Go to next year",
  closeButtonAriaLabel: "Close date picker"
};

export interface IDatePickerProps {
  inputDateChanged?: (newValue: Date) => void;
  inputDate?: Date;
  placeholder?: string;
}

export interface IDatePickerState
  extends React.ComponentState,
    IDatePickerProps {}

export class NLPDatePicker extends React.Component<
  IDatePickerProps,
  IDatePickerState
> {
  constructor(props: IDatePickerProps) {
    super(props);

    this.state = {
      inputDate: props.inputDate,
      placeholder: props.placeholder
    };
  }

  render(): JSX.Element {
    return (
      <div>
        <DatePicker
          strings={DayPickerStrings}
          firstWeekOfYear={1}
          placeholder={this.state.placeholder}
          ariaLabel={this.state.placeholder}
          allowTextInput={true}
          value={this.state.inputDate}
          formatDate={date =>
            `${date.toLocaleDateString(
              navigator.languages && navigator.languages[0]
            )}`
          }
          showWeekNumbers={false}
          onSelectDate={this.onSelectDate}
          parseDateFromString={this._onParseDateFromString}
        />
      </div>
    );
  }

  private onSelectDate = (value: Date): void => {
    this.setState({ inputDate: value });

    if (this.props.inputDateChanged) {
      this.props.inputDateChanged(value);
    }
  };

  private _onParseDateFromString = (value: string): Date => {

     let dateValue = Sugar.Date.create(value);
    return new Date(dateValue);
  };
}
