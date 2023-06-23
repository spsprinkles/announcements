import { Components, ContextInfo, Helper } from "gd-sprest-bs";
import { infoSquare } from "gd-sprest-bs/build/icons/svgs/infoSquare";
import { Datatable } from "./datatable";
import { DataSource, IListItem } from "./ds";
import { Security } from "./security";
import Strings from './strings';

/**
 * Banner
 */
export class Banner {
    private _dt: Datatable = null;
    private _el: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Save the properties
        this._el = el;

        // Render the banner
        this.render();
    }

    // Generates a link element
    private generateLinkElement(item: IListItem) {
        // Create the element
        let elItem = document.createElement("div");
        elItem.classList.add("announcement-item");
        elItem.classList.add("p-3");
        elItem.innerHTML = `<b>${item.Title}</b> ${item.Content}`;
        item.BackgroundColor ? elItem.style.backgroundColor = item.BackgroundColor : null;
        item.TextColor ? elItem.style.color = item.TextColor : null;

        // See if there is a link associated w/ this item
        if (item.LinkUrl && item.LinkUrl.Url) {
            // Add a click event
            elItem.addEventListener("click", () => {
                // Display in a new tab
                window.open(item.LinkUrl.Url, "_blank");
            });
        }

        // Return the element
        return elItem;
    }

    // Renders the banner
    private render() {
        // Create the datatable if it doesn't exist
        this._dt = this._dt || new Datatable(() => {
            // Clear the element
            while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

            // Render the component
            this.render();
        });

        // See if we are editing the page & in classic mode
        let isEditMode = Strings.IsClassic && Helper.WebPart.isEditMode();
        if (isEditMode) {
            // Render the edit button
            this.renderEdit();
        }

        // Ensure links exist
        if (DataSource.ListItems.length > 0) {
            // Render the dashboard
            this.renderAnnouncements();

            // Update the theme
            this.updateTheme();
        }
        // Ensure we aren't in edit mode and this is an owner/admin
        else if (!isEditMode && (Security.IsAdmin || Security.IsOwner)) {
            // Render the edit button
            this.renderEmptyBanner();
        }
    }

    // Renders the announcements
    private renderAnnouncements() {
        let elTicker = document.createElement("div");
        elTicker.classList.add("announcement-items");
        elTicker.classList.add("d-flex");
        elTicker.classList.add("flex-column");
        elTicker.classList.add("text-center");
        this._el.appendChild(elTicker);

        // Parse the items
        for (let i = 0; i < DataSource.ListItems.length; i++) {
            // Generate the element for this item
            let elItem = this.generateLinkElement(DataSource.ListItems[i]);

            // Append the element
            elTicker.appendChild(elItem);
        }
    }

    // Renders the edit information
    private renderEdit() {
        // Render a button to Edit Icon Links
        let btn = Components.Button({
            el: this._el,
            className: "ms-1 my-1",
            iconClassName: "btn-img",
            iconSize: 22,
            iconType: infoSquare,
            isSmall: true,
            text: "Edit Announcements",
            type: Components.ButtonTypes.OutlinePrimary,
            onClick: () => {
                // Show the datatable
                this._dt.show();
            }
        });

        // Update the style
        btn.el.classList.remove("btn-icon");
    }

    // Renders the empty banner
    private renderEmptyBanner() {
        let elTicker = document.createElement("div");
        elTicker.classList.add("announcement-items");
        this._el.appendChild(elTicker);

        // Render a button to Edit Icon Links
        let btn = Components.Button({
            el: this._el,
            className: "ms-1 my-1 announcement-item",
            iconClassName: "btn-img",
            iconSize: 22,
            iconType: infoSquare,
            isSmall: true,
            text: "Add Announcements",
            type: Components.ButtonTypes.OutlinePrimary,
            onClick: () => {
                // Show the datatable
                this._dt.show();
            }
        });

        // Update the style
        btn.el.classList.remove("btn-icon");
    }

    // Shows the datatable
    showDatatable() {
        // Show the datatable
        this._dt.show();
    }

    // Updates the styling, based on the theme
    updateTheme(themeInfo?: any) {
        // Get the theme colors
        let neutralDark = (themeInfo || ContextInfo.theme).neutralDark || DataSource.getThemeColor("StrongBodyText");
        let neutralLight = (themeInfo || ContextInfo.theme).neutralLight || DataSource.getThemeColor("DisabledLines");

        // Set the CSS properties to the theme colors
        let root = document.querySelector(':root') as HTMLElement;
        root.style.setProperty('--sp-neutral-dark', neutralDark);
        root.style.setProperty('--sp-neutral-light', neutralLight);
    }
}