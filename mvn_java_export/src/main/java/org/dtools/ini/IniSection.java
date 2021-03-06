/*
 * Java INI Package
 * Copyright (C) 2008 David Lewis
 * 
 * This program is free software: you can redistribute it and/or modify it under
 * the terms of the GNU Lesser General Public Licence as published by the Free
 * Software Foundation, either version 3 of the License, or (at your option) any
 * later version.
 * 
 * This program is distributed in the hope that it will be useful, but WITHOUT
 * ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 * FOR A PARTICULAR PURPOSE. See the GNU Lesser General Public Licence for more
 * details.
 * 
 * You should have received a copy of the GNU General Public License along with
 * this program. If not, see <http://www.gnu.org/licenses/>.
 */

package org.dtools.ini;

import java.util.ArrayList;
import java.util.Collection;

/**
 * <p>An IniSection represents a section within a configuration file (aka INI
 * file). A section can contain zero or more items (which have unique names
 * within a section), and those items have a value. In this API, an item is
 * represented by the class <code>IniItem</code>.</p>
 * 
 * @author David Lewis
 * @version 1.1.0
 * @since 0.1.10
 */
public abstract class IniSection implements Commentable, Cloneable, Iterable<IniItem> {
    
    /**
     * The name of this section
     */
    private String name;
    
    /**
     * Predicate that is true if this IniSection is case sensitive.
     */
    private boolean caseSensitive;
    
    /**
     * The comment that comes before the item
     */
    private String preComment;
    
    /**
     * The comment that comes after the item, but on the same line
     */
    private String endLineComment;
    
    /**
     * The comment that comes after the item.
     */
    private String postComment;

    /**
     * <p>A reference to this (and subclasses) <code>IniSection</code>'s 
     * <code>IniValidator</code> which validates the name of this <code>IniSection</code>
     * and all <code>IniItem</code>s within this <code>IniSection</code>.</p>
     */
    protected IniValidator validator;

    /**
     * <p>Default constructor which creates a new instance of this abstract
     * <code>IniSection</code> and sets the <code>IniSection</code> to have a
     * default <code>IniValidator</code> and to be case insensitive.</p>
     * 
     * @param name The name of this <code>IniSection</code>.
     * @throws InvalidNameException if the name of the <code>IniSection</code>
     * given is not considered valid by this object's <code>IniValidator</code>.
     */
    public IniSection( String name ) {
        this( name, new IniValidator() );
    }

    /**
     * <p>Constructor which creates a new instance of this abstract
     * <code>IniSection</code> and sets the <code>IniSection</code> to have a
     * default <code>IniValidator</code>.</p>
     * 
     * @param name The name of this <code>IniSection</code>.
     * @param caseSensitive Sets whether this IniSection is case sensitive or
     * not.
     * @throws InvalidNameException if the name of the <code>IniSection</code>
     * given is not considered valid by this object's <code>IniValidator</code>.
     */
    public IniSection(String name, boolean caseSensitive ) {
        this( name, new IniValidator(), caseSensitive );
    }

    /**
     * <p>Constructor which creates a new instance of this abstract
     * <code>IniSection</code> and sets the <code>IniSection</code> to be case
     * insensitive.</p>
     * 
     * @param name The name of this <code>IniSection</code>.
     * @param validator The <code>IniValidator</code> of this 
     * <code>IniSection</code>.
     * @throws InvalidNameException if the name of the <code>IniSection</code>
     * given is not considered valid by this object's <code>IniValidator</code>.
     */
    public IniSection(String name, IniValidator validator ) {
        this( name, validator, false );
    }

    /**
     * <p>Constructor which creates a new instance of this abstract
     * <code>IniSection</code>.</p>
     * 
     * @param  name The name of this <code>IniSection</code>.
     * @param  validator The <code>IniValidator</code> of this 
     * <code>IniSection</code>.
     * @param  caseSensitive Sets whether this IniSection is case sensitive or
     * not.
     * @throws InvalidNameException if the name of the <code>IniSection</code>
     * given is not considered valid by this object's <code>IniValidator</code>.
     */
    public IniSection(String name, IniValidator validator,
                                                       boolean caseSensitive ) {
        
        this.validator = validator;
        
        if( !validator.isValidSectionName(name) ) {
            throw new InvalidNameException( "The name, \"" + name + "\", is " +
                    "invalid for this section." );
        }
        
        this.caseSensitive = caseSensitive;
        
        this.name = name;
        
        setPreComment( "" );
        setEndLineComment( "" );
        setPostComment( "" );
    }
    
    /**
     * Adds an IniItem to this section.
     * 
     * If an item already in the section has the same name as the one being
     * added, then the method will not overwrite the item and will instead
     * return false.
     * 
     * The index of the item, if it added successfully, is unknown and depends
     * on the implementation of this interface. Programmers should not make any
     * assumptions on where new items are added. 
     *  
     * @param item the item to add
     * @return true if the item was added successfully, false if the item
     * already existed or could not be added
     */
    public boolean addItem( IniItem item ) {

        // cannot add null items
        if( item == null ) {
            return false;
        }

        String itemName = item.getName();

        //**********************************************************************
        // check that the item is compatible with this section
        //**********************************************************************
        if( !validator.isValidItemName(itemName) ) {
            throw new InvalidNameException( "The item's name, \"" + itemName + 
                    "\", is invalid for this section." );            
        }
        
        if( !validator.equals(item.getValidator()) ) {
            
            throw new InvalidNameException( "The item's name, \"" + 
                    itemName + "\", is valid for this section, but the " +
                    "item has an IniValidator that is incompatible with " +
                    "this section's IniValidator." );
        }

        //**********************************************************************
        // add the item
        //**********************************************************************
        if( hasItem(itemName) ) {
            return false;
        }
        else {
            return addItem( item, this.getNumberOfItems() );
        }
    }

    /**
     * Adds an IniItem to this section at the given index.
     * 
     * If an item already in the section has the same name as the one being
     * added, then the method will not overwrite the item and will instead
     * return false.
     * 
     * @param item the item to add
     * @param index the index where to add the item, where 0 is the index of the
     *        first item. Any items that already exists at this index will be
     *        moved to <code>index + 1</code>. If the value is greater than the
     *        number of items within this section, then the item is appended to
     *        the end of the section. 
     * @return true if the item was added successfully, false if the item
     *         already existed or could not be added
     * @throws IndexOutOfBoundsException if the value of <code>index</code> is
     *         less than 0.
     */
    public abstract boolean addItem( IniItem item, int index );

    /**
     * Adds a new IniItem to this section with the given name.
     * 
     * If an item already in the section has an the same name, then the
     * method returns false.
     * 
     * The index of the item, if it added successfully, is unknown and depends
     * on the implementation of this interface. Programmers should not make any
     * assumptions on where new items are added. 
     *  
     * @param itemName The name of the item to add.
     * @return A reference to the Item added to this section, or null if the
     * item could not be created.
     */
    public IniItem addItem( String itemName ) {
      
        // cannot add null items
        if( itemName == null ) {
            return null;
        }
        
        if( !validator.isValidItemName(itemName) ) {
            throw new InvalidNameException( "The item's name, \"" + itemName + 
                    "\", is invalid for this section." );            
        }
        
        if( hasItem(itemName) ) {
            return null;
        }
        
        else {
            IniItem newItem = createItem( itemName );
            addItem( newItem, this.getNumberOfItems() );
            return newItem;
        }
    }
    
    /**
     * Adds an IniItem to this section at the given index.
     * 
     * If an item already in the section has an the same name, then the
     * method returns false.
     * 
     * @param itemName The name of the item to add.
     * @param index The index where to add the item, where 0 is the index of the
     *        first item. Any items that already exists at this index will be
     *        moved to <code>index + 1</code>. If the value is greater than the
     *        number of items within this section, then the item is appended to
     *        the end of the section. 
     * @return A reference to the Item added to this section, or null if the
     *        item could not be created.
     * @throws IndexOutOfBoundsException if the value of <code>index</code> is
     *         less than 0.
     */
    public IniItem addItem( String itemName, int index ) {
      
        // cannot add null items
        if( itemName == null ) {
            return null;
        }
        
        if( !validator.isValidItemName(itemName) ) {
            throw new InvalidNameException( "The item's name, \"" + itemName + 
                    "\", is invalid for this section." );            
        }
        
        if( hasItem(itemName) ) {
            return null;
        }
        else {
            IniItem newItem = createItem( itemName );
            addItem( newItem, index );
            return newItem;
        }
    }
    
    /**
     * Add multiple new items to this section. The values of the new items will
     * be null. If any of the given names are duplicates, or if a name is
     * identical to an item that already exist.
     * 
     * The index of the item, if it added successfully, is unknown and depends
     * on the implementation of this interface. Programmers should not make any
     * assumptions on where new items are added. 
     * @param itemNames the names of the items to add.
     */
    public void addItems(String... itemNames) {
        
        for( String itemName : itemNames ) {
            addItem( createItem(itemName) );
        }
    }
    
    /**
     * </p>This method creates and returns a new instance of an 
     * <code>IniItem</code> with the same <code>IniValidator</code> and case
     * sensitivity as this object.</p> 
     * 
     * @param name The name of the <code>IniItem</code> to create.
     * @return A new instance of an <code>IniItem</code> with the same 
     * <code>IniValidator</code> and case sensitivity as this object.
     * @throws InvalidNameException if the name of the <code>IniItem</code>
     * given is not considered valid by this object's <code>IniValidator</code>.
     */
    protected abstract IniItem createItem( String name );
    
    /**
     * <p>This predicate returns true if this <code>IniSection</code> is equal
     * to the given object. For <code>other</code> to be equal to this one it
     * must:</p>
     * 
     * <ul>
     *   <li>be an instance of <code>IniSection</code>.</li>
     *   <li>have the same case-sensitivity as this <code>IniSection</code>.
     *   <li>have an equal <code>IniValidator</code> as this
     *       <code>IniSection</code>'s <code>IniValidator</code>.</li>
     *   <li>have an equal name as this <code>IniSection</code>.*</li>
     *   <li>have an equal pre-comment as this <code>IniSection</code>.*</li>
     *   <li>have an equal post-comment as this <code>IniSection</code>.*</li>
     *   <li>have an equal end-line-comment as this
     *       <code>IniSection</code>.*</li>
     *   <li>have the same number of <code>IniItem</code>s as this
     *       <code>IniSection</code>.</li>
     *   <li>have equal <code>IniItem</code>s as this
     *       <code>IniSection</code>.</li>
     *   <li>have the same order of <code>IniItem</code>s as this
     *       <code>IniSection</code>.</li>
     * </ul>
     * 
     * <p>* If the two <code>IniSection</code>s are
     * <strong>case-sensitive</strong>, then these are compared using the
     * <code>equals(Object)</code> method. If the two <code>IniSection</code>s
     * are <strong>case-insensitive</strong>, then these are compared using the
     * <code>String.equalsIgnoreCase(String)</code> method.</p>
     * 
     * @param other The other Object to test for equality.
     * @return True if equal, false if not equal.
     * @since 0.1.15
     */
    @Override
    public boolean equals( Object other ) {
        
        IniSection otherSection;

        //**********************************************************************
        // check to see if the other section is an instance of the same class
        //**********************************************************************
        if( !(other instanceof IniSection) ) {
            return false;
        } else {
            otherSection = (IniSection) other;
        }

        //**********************************************************************
        // check that the two sections have the same case sensitive
        //**********************************************************************
        if( this.isCaseSensitive() != otherSection.isCaseSensitive() ) {
            return false;
        }

        //**********************************************************************
        // check that the two sections have the same name
        //**********************************************************************
        if( !(otherSection.getName().equals(this.getName())) ) {
            return false;
        }

        //**********************************************************************
        // check to see if they have the same validator
        //**********************************************************************
        if( !this.getValidator().equals(otherSection.getValidator()) ) {
            return false;
        }
        
        //**********************************************************************
        // check to see if the two sections have the same number of items
        //**********************************************************************
        // 
        if( this.getNumberOfItems() != otherSection.getNumberOfItems() ) {
            return false;
        }

        //**********************************************************************
        // check that the sections have the end line comments
        //**********************************************************************
        String thisComment = this.getEndLineComment();
        String otherComment = otherSection.getEndLineComment();
        
        if( !thisComment.equals(otherComment) ) {
            return false;
        }

        //**********************************************************************
        // check to see if the two sections have the same pre comments
        //**********************************************************************
        thisComment = this.getPreComment();
        otherComment = otherSection.getPreComment();
        
        if( !thisComment.equals(otherComment) ) {
            return false;
        }

        //**********************************************************************
        // check to see if the two sections have the same post comments
        //**********************************************************************
        thisComment = this.getPostComment();
        otherComment = otherSection.getPostComment();
            
        if( !thisComment.equals(otherComment) ) {
            return false;
        }

        //**********************************************************************
        // check to see if the two sections have the same items
        //**********************************************************************
        for( int i=0; i<this.getNumberOfItems(); i++ ) {
            
            IniItem thisItem = this.getItem( i );
            IniItem otherItem = otherSection.getItem( i );
            
            // if this item does not equal the other item then return false
            if( !thisItem.equals(otherItem) ) {
                return false;
            }
        }
        
        // the two sections are equal by method of deduction
        return true;
    }

    @Override
    public String getEndLineComment() {
        return endLineComment;
    }

    /**
     * Get the item which is at the given index. 
     * 
     * @param index the index of the item to retrieve.
     * @return The item that is at the given index.
     * @throws IndexOutOfBoundsException if the given value is less than 0 or
     *         greater or equal to the number of items in this section (i.e.
     *         <code>&gt; getNumberOfItems()-1</code>.
     */
    public abstract IniItem getItem( int index );

    /**
     * Returns the item with the given name, or null if no item exists with the 
     * given name.
     * 
     * @param name The name of the item.
     * @return The IniItem with the given name, or null if the item doesn't
     * exists.
     */
    public IniItem getItem(String name) {
        
        for( IniItem item : getItems() ) {
            
            if( caseSensitive ) {
                if( item.getName().equals(name) ) {
                    return item;
                }
            }
            else {
                if( item.getName().equalsIgnoreCase(name) ) {
                    return item;
                }
            }
        }
        
        return null;
    }

    /**
     * This method returns a collection of all the names of all the items within
     * this section.
     * 
     * @return A collection of all the names.
     */
    public Collection<String> getItemNames() {
        
        Collection<String> itemNames = new ArrayList<String>();
        
        for( IniItem item : getItems() ) {
            itemNames.add( item.getName() );
        }
        
        return itemNames;
    }

    /**
     * Get a collection of all the items within this section.
     * 
     * @return A collection of all the items.
     */
    public abstract Collection<IniItem> getItems();

    /**
     * This method returns the name of this section. Note that a section cannot
     * change its name. 
     * 
     * @return The name of this section 
     */
    public String getName() {
        return name;
    }

    /**
     * Return the number of IniItems that this array has.
     * 
     * @return the number of items that this array has
     */
    public int getNumberOfItems() {
        return getItems().size();
    }

    @Override
    public String getPostComment() {
        return postComment;
    }

    @Override
    public String getPreComment() {
        return preComment;
    }

    /**
     * <p>Gets a reference to this object's <code>IniValidator</code>.</p>
     * 
     * @return The <code>IniValidator</code> used by this
     * <code>IniSection</code>.
     * @since 0.1.17
     */
    public IniValidator getValidator() {
        return validator;
    }


    @Override
    public int hashCode() {
        
        int total = getName().hashCode();
        
        for( IniItem item : getItems() ) {
            total = (total + item.hashCode()) % Integer.MAX_VALUE;
        }
        
        return total;
    }
    
    /**
     * Predicate that returns true if an item is in this section. More
     * specifically this method returns true if the class contains a reference
     * to the given item, and not if this class has a similar item with
     * the same name and value.
     * 
     * @param item The item to test.
     * @return True if an item that matches that name exists, false otherwise.
     */
    public boolean hasItem(IniItem item) {
        
        for( IniItem i : getItems() ) {
            
            if( i == item ) {
                return true;
            }
        }
        
        // this section does not contain the given Item
        return false;
    }
    
    /**
     * Predicate that returns true if an item already exists.
     * 
     * @param itemName the name of the item to test
     * @return true if an item that matches that name exists, false otherwise.
     */
    public boolean hasItem( String itemName ) {
        
        for( IniItem item : getItems() ) {
            
            String name = item.getName();
            
            if( caseSensitive ) {
                if( name.equals(itemName) ) {
                    return true;
                }
            }
            else {
                if( name.equalsIgnoreCase(itemName) ) {
                    return true;
                }
            }
        }
        
        // this section does not contain the given Item
        return false;
    }

    /**
     * Returns the index of item that is in this section. If no such item
     * exists, then the value <code>-1</code> is returned.
     * 
     * @param item The IniItem whose index will be returned.
     * @return The index of the given item.
     */
    public abstract int indexOf( IniItem item );

    /**
     * Returns the index of item that is in this section. If no such item
     * exists, then the value <code>-1</code> is returned.
     * 
     * @param itemName The name of the item whose index will be returned.
     * @return The index of the item.
     */
    public int indexOf(String itemName) {
        return indexOf( getItem(itemName) );
    }

    /**
     * <p>Predicate that returns true if this <code>IniSection</code> is case
     * <strong>sensitive</strong>, or false if this <code>IniSection</code> is
     * case <strong>insensitive</strong>.</p>
     * 
     * @return boolean
     * @since 0.1.16
     */
    public boolean isCaseSensitive() {
        return caseSensitive;
    }
    
    /**
     * Predicate that returns true if the section has no items, false if it does
     * have at least one item.
     * 
     * @return true or false
     */
    public boolean isEmpty() {
        return getItems().isEmpty();
    }
    
    /**
     * This method merges another IniSection to this section. All the items
     * that the other section has will be copied over to this section providing
     * that no item has the same name in both section. 
     * 
     * @param otherSection the other section to merge to this section
     * @return true if the merge was successful, false if the merge failed due
     * to an item that had the same name in both sections.
     */
    public boolean merge( IniSection otherSection ) {

        //**********************************************************************
        // Step 1 - check if a merge is possible
        // 
        // NOTE: if an item, with the same name exists in this and the given
        // IniSection, then a merge is not possible.
        //**********************************************************************
        if( !IniUtilities.isDisjoint(this,otherSection) ) {
            return false;
        }
        
        // check to see if they both have the same validator
        if( !this.validator.equals(otherSection.getValidator()) ) {
            return false;
        }
        
        if( this.isCaseSensitive() != otherSection.isCaseSensitive() ) {
            return false;
        }
        
        //**********************************************************************
        // Step 2 - merge is possible, perform merge
        //**********************************************************************
        for( IniItem otherItem : otherSection.getItems() ) {
            this.addItem( otherItem );
        }
        
        return true;
    }
    
    /**
     * This method moves an item at index <code>fromIndex</code> to the index
     * <code>toIndex</code>. 
     * 
     * @param fromIndex The index of the item to move
     * @param toIndex The index where to place the item, any items already at
     *        that index will be moved to <code>index+1</code>.
     * @throws IndexOutOfBoundsException if either <code>fromIndex</code> or
     *         <code>toIndex</code> is below 0 or greater than or equal to the
     *         number of items in this section (i.e.
     *         <code>&gt; getNumberOfItems()-1</code>).
     */
    public void moveItem(int fromIndex, int toIndex) {

        if( fromIndex < 0 || fromIndex >= this.getNumberOfItems() ) {
            throw new IndexOutOfBoundsException( "" + fromIndex );
        }
        
        if( toIndex < 0 || toIndex >= this.getNumberOfItems() ) {
            throw new IndexOutOfBoundsException( "" + toIndex );
        }
        
        IniItem section = getItem( fromIndex );
        removeItem( section );
        addItem( section, toIndex );
    }
    
    /**
     * This method moves the item whose names is given to the the index
     * <code>toIndex</code>.
     * 
     * @param name The name of the item to move.
     * @param toIndex The index where to place the item, any items already at
     *        that index will be moved to <code>index+1</code>.
     * @throws IndexOutOfBoundsException if <code>toIndex</code> is below 0 or
     *         greater than or equal to the number of items in this section
     *         (i.e. <code>&gt; getNumberOfItems()-1</code>), or if the given 
     *         name of the item doesn't exist.
     */
    public void moveItem(String name, int toIndex) {
        moveItem( indexOf(name), toIndex);
    }

    /**
     * Removes all the items and its values from the section.
     */
    public void removeAll() {
        
        for( IniItem item : getItems() ) {
            removeItem( item );
        }
    }

    @Override
    public void removeEndLineComment() {
        setEndLineComment("");
    }

    /**
     * Removes the item this section. If the given item does not exists in the
     * section, then the method does nothing.
     * 
     * @param item The item to remove.
     * @return boolean, true if the IniItem was removed, false it the given
     * IniItem did not exists in the first place. 
     */
    public abstract boolean removeItem( IniItem item );

    /**
     * Removes the item that is at the given index from this section.
     * 
     * @param index the index of the item to remove
     * @return The IniItem that was removed.
     * @throws IndexOutOfBoundsException if <code>index</code> is below 0 or
     * greater than or equal to the number of items in this section (i.e. 
     * <code>&gt; getNumberOfItems() - 1</code>).
     */
    public IniItem removeItem( int index ) {

        IniItem item = getItem( index );
        removeItem( item );
        return item;
    }

    /**
     * Removes an item whose name matches the given name the section.
     * 
     * @param name The name of the name of the item to remove.
     * @return The IniItem that was removed.
     */
    public IniItem removeItem(String name) {
        IniItem item = getItem( name );
        removeItem( item );
        return item;
    }

    
    /**
     * Removes all the items from this section whose names are within the given
     * array.
     * 
     * @param names An array of the names of the items to remove
     * @return A collection of all the IniItems that were removed.
     */
    public Collection<IniItem> removeItems(String[] names) {

        // create a new collection
        Collection<IniItem> results = new ArrayList<IniItem>(names.length);
        
        for( String name : names ) {
            
            IniItem item = removeItem( name );
            
            if( item != null ) {
                results.add( item );
            }
        }
        
        return results;
    }

    @Override
    public void removePostComment() {
        setPostComment("");
    }

    @Override
    public void removePreComment() {
        setPreComment("");
    }

    @Override
    public void setEndLineComment(String comment) {

        // if comment is null, set to be empty string
        if( comment == null ) {
            comment = "";
        }
        
        // check that comment doesn't contain a new line character
        if( comment.contains("\n") ) {
            throw new IllegalArgumentException( "A comment cannot contain a " +
                    "new line character for an end line comment." );
        }
        
        if( !isCaseSensitive() ) {
            comment = comment.toLowerCase();
        }
        
        endLineComment = comment;
    }

    @Override
    public void setPostComment(String comment) {
        
        if( comment == null ) {
            comment = "";
        }
        
        if( !isCaseSensitive() ) {
            comment = comment.toLowerCase();
        }
        
        postComment = comment;
    }
    
    @Override
    public void setPreComment(String comment) {
        
        if( comment == null ) {
            comment = "";
        }
        
        if( !isCaseSensitive() ) {
            comment = comment.toLowerCase();
        }
        
        preComment = comment;
    }
    
    @Override
    public String toString() {
        
        StringBuilder sb = new StringBuilder();
        
        sb.append( "org.dtools.ini.IniSection \"" + getName() + "\": " );
        sb.append( "(Items: " + getNumberOfItems() + ")" );
        
        for( IniItem item : getItems() ) {
            sb.append( "\n" );
            sb.append( item.toString() );
        }
        
        return sb.toString();
    }
    
    /**
     * <p>This class implements the Cloneable interface, but leaves
     * implementation to subclasses which override this class.
     * 
     * @return An object which is an instance of class, has the same values as
     * this object, would return true for <code>thisObject.equals(
     * thisObject.clone() );</code>, but returns false for <code>thisObject ==
     * thisObject.clone();</code>
     * 
     * @since 1.0.00
     */
    @Override
    public abstract Object clone();
}
