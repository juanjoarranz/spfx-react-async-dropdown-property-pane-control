declare interface IListItemsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName         : string;
  ListFieldLabel         : string;
  ItemFieldLabel         : string;
}

declare module 'ListItemsWebPartStrings' {
  const strings: IListItemsWebPartStrings;
  export = strings;
}
