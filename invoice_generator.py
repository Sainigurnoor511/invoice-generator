from tkinter import Tk, Frame, ttk, messagebox, END
from docxtpl import DocxTemplate
import datetime
import os
import re


class InvoiceGenerator:
    """Invoice Generator application using Tkinter GUI."""
    
    # Constants
    WINDOW_WIDTH = 900
    WINDOW_HEIGHT = 460
    SALES_TAX_RATE = 0.1
    TEMPLATE_FILE = "invoice_template.docx"
    INVOICES_FOLDER = "generated_invoices"
    
    def __init__(self):
        """Initialize the Invoice Generator application."""
        self.root = Tk()
        self.root.title("Invoice Generator Form")
        self.invoice_list = []
        
        # Center window on screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_coordinate = (screen_width / 2) - (self.WINDOW_WIDTH / 2)
        y_coordinate = (screen_height / 2) - (self.WINDOW_HEIGHT / 1.8)
        self.root.geometry(f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}+{int(x_coordinate)}+{int(y_coordinate)}")
        self.root.resizable(width=False, height=False)

    def widgets(self):
        """Create and layout all GUI widgets."""
        self.frame = Frame(self.root)
        self.frame.pack(padx=20, pady=10)

        # Customer information section
        self.first_name_label = ttk.Label(self.frame, text="First Name *")
        self.first_name_label.grid(row=0, column=0, sticky="w", padx=5)
        self.last_name_label = ttk.Label(self.frame, text="Last Name *")
        self.last_name_label.grid(row=0, column=1, sticky="w", padx=5)

        self.first_name_entry = ttk.Entry(self.frame, width=30)
        self.last_name_entry = ttk.Entry(self.frame, width=30)
        self.first_name_entry.grid(row=1, column=0, padx=5)
        self.last_name_entry.grid(row=1, column=1, padx=5)

        self.phone_label = ttk.Label(self.frame, text="Phone *")
        self.phone_label.grid(row=0, column=2, sticky="w", padx=5)
        self.phone_entry = ttk.Entry(self.frame, width=30)
        self.phone_entry.grid(row=1, column=2, padx=5)

        # Item details section
        self.qty_label = ttk.Label(self.frame, text="Quantity *")
        self.qty_label.grid(row=2, column=0, sticky="w", padx=5, pady=(10, 0))
        self.qty_spinbox = ttk.Spinbox(self.frame, from_=1, to=100, width=28)
        self.qty_spinbox.set(1)
        self.qty_spinbox.grid(row=3, column=0, padx=5)

        self.desc_label = ttk.Label(self.frame, text="Description *")
        self.desc_label.grid(row=2, column=1, sticky="w", padx=5, pady=(10, 0))
        self.desc_entry = ttk.Entry(self.frame, width=30)
        self.desc_entry.grid(row=3, column=1, padx=5)

        self.price_label = ttk.Label(self.frame, text="Unit Price *")
        self.price_label.grid(row=2, column=2, sticky="w", padx=5, pady=(10, 0))
        self.price_spinbox = ttk.Spinbox(self.frame, from_=0.0, to=10000, increment=0.5, width=28)
        self.price_spinbox.set(0.0)
        self.price_spinbox.grid(row=3, column=2, padx=5)

        self.add_item_button = ttk.Button(self.frame, text="Add Item", command=self.add_item)
        self.add_item_button.grid(row=4, column=2, pady=5, padx=5, sticky="e")

        # Invoice items treeview
        self.columns = ('qty', 'desc', 'price', 'total')
        self.tree = ttk.Treeview(self.frame, columns=self.columns, show="headings", height=6)
        self.tree.heading('qty', text='Quantity')
        self.tree.heading('desc', text='Description')
        self.tree.heading('price', text='Unit Price')
        self.tree.heading('total', text="Total")
        
        # Set column widths
        self.tree.column('qty', width=80, anchor='center')
        self.tree.column('desc', width=300)
        self.tree.column('price', width=100, anchor='e')
        self.tree.column('total', width=100, anchor='e')
            
        self.tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)

        # Action buttons
        self.save_invoice_button = ttk.Button(
            self.frame, text="Generate Invoice", command=self.generate_invoice
        )
        self.save_invoice_button.grid(row=6, column=0, columnspan=3, sticky="ew", padx=20, pady=5)
        
        self.new_invoice_button = ttk.Button(
            self.frame, text="New Invoice", command=self.new_invoice
        )
        self.new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="ew", padx=20, pady=5)

    def validate_customer_info(self):
        """Validate customer information fields."""
        first_name = self.first_name_entry.get().strip()
        last_name = self.last_name_entry.get().strip()
        phone = self.phone_entry.get().strip()
        
        if not first_name:
            messagebox.showerror("Validation Error", "First name is required!")
            self.first_name_entry.focus()
            return False
        
        if not last_name:
            messagebox.showerror("Validation Error", "Last name is required!")
            self.last_name_entry.focus()
            return False
        
        if not phone:
            messagebox.showerror("Validation Error", "Phone number is required!")
            self.phone_entry.focus()
            return False
        
        # Basic phone validation (at least 10 digits)
        phone_digits = re.sub(r'\D', '', phone)
        if len(phone_digits) < 10:
            messagebox.showerror("Validation Error", "Please enter a valid phone number (at least 10 digits)!")
            self.phone_entry.focus()
            return False
        
        return True
    
    def validate_item_fields(self):
        """Validate item input fields before adding to invoice."""
        try:
            qty = int(self.qty_spinbox.get())
            if qty <= 0:
                messagebox.showerror("Validation Error", "Quantity must be greater than 0!")
                self.qty_spinbox.focus()
                return False
        except ValueError:
            messagebox.showerror("Validation Error", "Please enter a valid quantity!")
            self.qty_spinbox.focus()
            return False
        
        description = self.desc_entry.get().strip()
        if not description:
            messagebox.showerror("Validation Error", "Description is required!")
            self.desc_entry.focus()
            return False
        
        try:
            price = float(self.price_spinbox.get())
            if price <= 0:
                messagebox.showerror("Validation Error", "Unit price must be greater than 0!")
                self.price_spinbox.focus()
                return False
        except ValueError:
            messagebox.showerror("Validation Error", "Please enter a valid unit price!")
            self.price_spinbox.focus()
            return False
        
        return True
    
    def sanitize_filename(self, filename):
        """Remove invalid characters from filename."""
        # Remove invalid filename characters
        filename = re.sub(r'[<>:"/\\|?*]', '', filename)
        # Replace spaces with underscores
        filename = filename.replace(' ', '_')
        return filename

    def clear_item(self):
        """Clear item input fields."""
        self.qty_spinbox.delete(0, END)
        self.qty_spinbox.insert(0, "1")
        self.desc_entry.delete(0, END)
        self.price_spinbox.delete(0, END)
        self.price_spinbox.insert(0, "0.0")

    def add_item(self):
        """Add an item to the invoice list after validation."""
        # Validate customer info first
        if not self.validate_customer_info():
            return
        
        # Then validate item fields
        if not self.validate_item_fields():
            return
        
        try:
            qty = int(self.qty_spinbox.get())
            desc = self.desc_entry.get().strip()
            price = float(self.price_spinbox.get())
            line_total = qty * price
            
            invoice_item = [qty, desc, f"${price:.2f}", f"${line_total:.2f}"]
            self.tree.insert('', 0, values=invoice_item)
            
            # Store raw values for calculation
            self.invoice_list.append([qty, desc, price, line_total])
            
            self.clear_item()
            self.desc_entry.focus()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add item: {str(e)}")

    def new_invoice(self):
        """Clear all fields and start a new invoice."""
        self.first_name_entry.delete(0, END)
        self.last_name_entry.delete(0, END)
        self.phone_entry.delete(0, END)
        self.clear_item()
        self.tree.delete(*self.tree.get_children())
        self.invoice_list.clear()
        self.first_name_entry.focus()
        
    def generate_invoice(self):
        """Generate and save the invoice document."""
        # Validate customer information
        if not self.validate_customer_info():
            return
        
        # Check if invoice has items
        if not self.invoice_list:
            messagebox.showerror("Validation Error", "Please add at least one item to the invoice!")
            return
        
        # Check if template file exists
        if not os.path.exists(self.TEMPLATE_FILE):
            messagebox.showerror(
                "Template Error", 
                f"Template file '{self.TEMPLATE_FILE}' not found!\nPlease ensure the template file is in the same directory as this script."
            )
            return
        
        try:
            # Create invoices folder if it doesn't exist
            if not os.path.exists(self.INVOICES_FOLDER):
                os.makedirs(self.INVOICES_FOLDER)
            
            doc = DocxTemplate(self.TEMPLATE_FILE)
            
            # Get customer information
            first_name = self.first_name_entry.get().strip()
            last_name = self.last_name_entry.get().strip()
            name = f"{first_name} {last_name}"
            phone = self.phone_entry.get().strip()
            
            # Calculate totals
            subtotal = sum(item[3] for item in self.invoice_list)
            sales_tax = subtotal * self.SALES_TAX_RATE
            total = subtotal + sales_tax
            
            # Render the document
            doc.render({
                "name": name,
                "phone": phone,
                "invoice_list": self.invoice_list,
                "subtotal": f"${subtotal:.2f}",
                "salestax": f"{self.SALES_TAX_RATE * 100:.0f}%",
                "tax_amount": f"${sales_tax:.2f}",
                "total": f"${total:.2f}"
            })
            
            # Generate sanitized filename
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
            sanitized_name = self.sanitize_filename(f"{first_name}_{last_name}")
            filename = f"invoice_{sanitized_name}_{timestamp}.docx"
            
            # Save to invoices folder
            doc_path = os.path.join(self.INVOICES_FOLDER, filename)
            doc.save(doc_path)
            
            messagebox.showinfo(
                "Invoice Complete", 
                f"Invoice generated successfully!\nSaved to: {doc_path}"
            )
            
            self.new_invoice()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate invoice:\n{str(e)}")

if __name__ == "__main__":
    app = InvoiceGenerator()
    app.widgets()
    app.first_name_entry.focus()
    app.root.mainloop()