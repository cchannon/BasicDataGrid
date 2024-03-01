import * as React from "react";
import {
  DetailsList,
  IColumn,
  Stack,
  SelectionMode,
  ConstrainMode,
  ColumnActionsMode,
  Text,
  StackItem,
  ISelection,
  Selection
} from "@fluentui/react";

export interface IGridProps {
  gridHeight: string;
  dataset: ComponentFramework.PropertyTypes.DataSet;
  navigate: (
    options: ComponentFramework.NavigationApi.EntityFormOptions
  ) => void;
  setSelected: (ids: string[]) => void;
}

export const DataGrid: React.FunctionComponent<IGridProps> = (
  props: IGridProps
) => {
  const [items, setItems] = React.useState<any[]>([]);
  const [columns, setColumns] = React.useState<IColumn[]>([]);
  const [sortFlag, setSortFlag] = React.useState<number>(0);
  const [sortColumn, setSortColumn] = React.useState<IColumn>();

  //Selection object for handling selection change events
  const selection: ISelection = new Selection({
    onSelectionChanged: () => {
      props.setSelected(selection.getSelection().map((item) => {
        return (item as any).id
      }));
    },
  });

  //set items if dataset changes
  React.useEffect(() => {
    setItems(props.dataset.sortedRecordIds.map((id: string) => {
      const item = props.dataset.records[id];
      return props.dataset.columns.reduce((acc, col) => {
        let val: any = null;
        if (
          col.dataType === "Lookup.Simple" ||
          col.dataType === "Lookup.Owner"
        ) {
          val = item.getValue(col.name);
        } else {
          val = item.getFormattedValue(col.name);
        }
        return {
          ...acc,
          id: item.getRecordId(),
          logicalName: (item.getNamedReference() as any).entityName,
          [col.name]: val,
        };
      }, {});
    }));
  }, [props.dataset.records]);

  //render column definitions only if changed
  React.useEffect(() => {
    let cols = props.dataset.columns
    .filter((col) => !col.isHidden)
    .sort((a, b) => a.order - b.order)
    .map((col: ComponentFramework.PropertyHelper.DataSetApi.Column) => {
      let column: IColumn = {
        key: `column_${col.name}`,
        name: col.displayName,
        fieldName: col.name,
        targetWidthProportion: col.visualSizeFactor,
        minWidth: 100,
        isResizable: true,
        columnActionsMode: ColumnActionsMode.clickable,
        onColumnClick: () => { setSortColumn(column); setSortFlag(Math.random()); },
        isSorted: false,
        isSortedDescending: false,
      };

      switch (col.dataType) {
        case "Lookup.Simple":
        case "Lookup.Owner":
          column.onRender = (item: any) => {
            let cval: ComponentFramework.EntityReference =
              item[column.fieldName!];
            if (cval) {
              return (
                <div
                  style={{
                    fontFamily:
                      '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                    fontSize: "14px",
                    fontWeight: "normal",
                    color: "rgb(17, 94, 163)",
                    textDecoration: "underline",
                    cursor: "pointer",
                  }}
                  onClick={() =>
                    props.navigate({
                      entityName: cval.etn!,
                      entityId: cval.id.guid,
                      openInNewWindow: true,
                      useQuickCreateForm: false,
                    })
                  }
                >
                  {cval.name ?? ""}
                </div>
              );
            } else {
              return <div></div>;
            }
          };
          break;
        case "SingleLine.Email":
          column.onRender = (item: any) => {
            const columnValue = item[column.fieldName!];
            return (
              <a
                href={`mailto:${columnValue ?? ""}`}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  fontFamily:
                    '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                  fontSize: "14px",
                  fontWeight: "normal",
                  color: "rgb(17, 94, 163)",
                }}
              >
                {columnValue ?? ""}
              </a>
            );
          };
          break;
        case "SingleLine.URL":
          column.onRender = (item: any) => {
            const columnValue = item[column.fieldName!];
            return (
              <a
                href={columnValue ?? ""}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  fontFamily:
                    '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                  fontSize: "14px",
                  fontWeight: "normal",
                  color: "rgb(17, 94, 163)",
                }}
              >
                {columnValue ?? ""}
              </a>
            );
          };
          break;
        case "Decimal":
        case "Currency":
        case "Whole.None":
        case "FP":
          column.onRender = (item: any) => {
            const columnValue = item[column.fieldName!];
            return (
              <div
                style={{
                  textAlign: "right",
                  fontFamily:
                    '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                  fontSize: "14px",
                  fontWeight: "normal",
                }}
              >
                {columnValue ?? ""}
              </div>
            );
          };
          break;
        default:
          if (
            props.dataset.columns.find((col) => col.name === column.fieldName)
              ?.isPrimary
          ) {
            column.onRender = (item: any) => {
              let cval = item[column.fieldName!];
              if (cval) {
                return (
                  <div
                    style={{
                      fontFamily:
                        '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                      fontSize: "14px",
                      fontWeight: "normal",
                      color: "rgb(17, 94, 163)",
                      textDecoration: "underline",
                      cursor: "pointer",
                    }}
                    onClick={() =>
                      props.navigate({
                        entityName: item.logicalName,
                        entityId: item.id,
                        openInNewWindow: true,
                        useQuickCreateForm: false,
                      })
                    }
                  >
                    {cval ?? ""}
                  </div>
                );
              } else {
                return <div></div>;
              }
            };
          } else {
            column.onRender = (item: any) => {
              const columnValue = item[column.fieldName!];
              return (
                <div
                  style={{
                    textAlign: "left",
                    fontFamily:
                      '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                    fontSize: "14px",
                    fontWeight: "normal",
                  }}
                >
                  {columnValue ?? ""}
                </div>
              );
            };
          }
      }
      return column;
    });
    
    setColumns(cols);
  }, [props.dataset.columns]);

  //sort items if sortColumn changes
  React.useEffect(() => {
    //set isSorted and isSortedDescending on columns
    if(!sortColumn) return;
    let currentColumn = columns.find((c) => c.key === sortColumn.key);
    let updatedColumns = [...columns].map((col: IColumn) => {
      if (col.fieldName === sortColumn.fieldName) {
        if (columns.find(c => c.key === sortColumn.key)!.isSorted)
          return {
            ...col,
            isSorted: true,
            isSortedDescending: !currentColumn!.isSortedDescending,
          };
        else
          return {
            ...col,
            isSorted: true,
            isSortedDescending: false,
          };
      } else if (col.isSorted) {
        return {
          ...col,
          isSorted: false,
          isSortedDescending: false,
        };
      } else return col;
    });
    setColumns(updatedColumns);

    const updatedColumn = updatedColumns.find((c) => c.key === sortColumn.key);

    //arrange items
    let updatedItems = [...items].sort((a: any, b: any) => {
        //sort based on data type
      if (typeof a[sortColumn.fieldName!] === "string") {
        return (
          (a[sortColumn.fieldName!] ?? "").localeCompare(
            b[sortColumn.fieldName!] ?? ""
          ) * (updatedColumn!.isSortedDescending ? -1 : 1)
        );
      } else if (typeof a[sortColumn.fieldName!] === "number") {
        return (
          (a[sortColumn.fieldName!] ?? 0 - b[sortColumn.fieldName!] ?? 0) *
          (updatedColumn!.isSortedDescending ? -1 : 1)
        );
      } else if (typeof a[sortColumn.fieldName!] === "boolean") {
        return (
          (a[sortColumn.fieldName!] === b[sortColumn.fieldName!]
            ? 0
            : a[sortColumn.fieldName!]
            ? 1
            : -1) * (updatedColumn!.isSortedDescending ? -1 : 1)
        );
      } else {
        return (
          (a[sortColumn.fieldName!]?.name ?? "").localeCompare(
            b[sortColumn.fieldName!]?.name ?? ""
          ) * (updatedColumn!.isSortedDescending ? -1 : 1)
        );
      }
    });
    setItems(updatedItems);
  }, [sortFlag]);

  function onItemInvoked(
    item?: any,
    _i?: number | undefined,
    _ev?: Event | undefined
  ): void {
    props.navigate({
      entityName: item.logicalName,
      entityId: item.id,
      openInNewWindow: true,
      useQuickCreateForm: false,
    });
  }

  return (
    <Stack
      tokens={{ childrenGap: 10, maxWidth: "100%" }}
      style={{ width: "100%" }}
    >
      <StackItem style={{ height: props.gridHeight }}>
        <DetailsList
          items={items}
          columns={columns}
          selectionMode={SelectionMode.none}
          selection={selection}
          constrainMode={ConstrainMode.horizontalConstrained}
          onItemInvoked={onItemInvoked}
          styles={{
            root: {
              overflowX: "scroll",
              overflowY: "scroll",
              height: props.gridHeight,
            },
          }}
        />
      </StackItem>
      <Text
        variant="smallPlus"
        style={{ textAlign: "left", marginLeft: "10px" }}
      >
        Total records found: {props.dataset.paging.totalResultCount}
      </Text>
    </Stack>
  );
};
