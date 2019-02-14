import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';

export interface IAsyncDropdownProps {
  label      : string;
  loadOptions: () => Promise<IDropdownOption[]>; // called by the control to load the available options.
  onChanged  : ( option: IDropdownOption, index?: number ) => void; //  called after the user selects an option in the dropdown.
  selectedKey: string | number; //specifies the selected value
  disabled   : boolean; //specifies if the dropdown control is disabled or not.
  stateKey   : string; //to force the React component to re-render.
}
