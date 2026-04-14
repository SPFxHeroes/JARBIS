/**
 * A record from the "Powers" SharePoint list.
 *
 * Represents a single record retrieved from the "Powers" SharePoint list, with properties for each of the columns in the list.
 */
export interface IPowerItem {
  /**
   * The title of the power.
   *
   * This is the title of the record in the SharePoint list.
   */
  Title: string;

  /**
   * The supported colors for the power.
   *
   * This is the list of colors that will work well with this power.
   */
  Colors: string[];

  /**
   * The icons for the power.
   *
   * This is the list of potential icons that can be used for this power.
   */
  Icon: string[];

  /**
   * The main text for the power.
   *
   * Contains a list of potential main text for the power.
   */
  Main: string[];

  /**
   * The prefix for the power.
   *
   * Contains a list of potential prefixes (e.g. Captain, Doctor, etc.) for the power.
   */
  Prefix: string[];
}
