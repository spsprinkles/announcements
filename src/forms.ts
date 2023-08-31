import { CanvasForm } from "dattatable";
import { DataSource, IListItem } from "./ds";

/**
 * Forms
 */
export class Forms {
    // Displays the edit form
    static edit(itemId: number, onUpdate: () => void) {
        // Update the canvas properties
        CanvasForm.clear();
        CanvasForm.setAutoClose(false);

        // Display the edit form
        DataSource.List.editForm({
            itemId,
            onSetFooter: (el) => {
                let btn = el.querySelector("button") as HTMLButtonElement;
                if (btn) { btn.innerText = "Save" }
            },
            onSetHeader: (el) => {
                let h5 = el.querySelector("h5") as HTMLElement;
                if (h5) { h5.prepend("Edit: "); }
            },
            onUpdate: (item: IListItem) => {
                // Refresh the item
                DataSource.List.refreshItem(item.Id).then(() => {
                    // Call the update event
                    onUpdate();
                });
            }
        });
    }

    // Displays the new form
    static new(onUpdate: () => void) {
        // Update the canvas properties
        CanvasForm.clear();
        CanvasForm.setAutoClose(false);

        // Display the new form
        DataSource.List.newForm({
            onSetHeader: (el) => {
                let h5 = el.querySelector("h5") as HTMLElement;
                if (h5) { h5.innerText = "New Announcement"; }
            },
            onUpdate: (item: IListItem) => {
                // Refresh the data
                DataSource.List.refreshItem(item.Id).then(() => {
                    // Call the update event
                    onUpdate();
                });
            }
        });
    }

    // Displays the view form
    static view(item: IListItem) {
        // Update the canvas properties
        CanvasForm.clear();
        CanvasForm.setAutoClose(false);

        // Display the view form
        DataSource.List.viewForm({ itemId: item.Id });
    }
}