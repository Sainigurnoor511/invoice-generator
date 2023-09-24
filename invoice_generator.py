import tkinter
from tkinter import *
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox

class InvoiceGenerator:
    def __init__(self):
        self.root = Tk()
        self.root.title("Invoice Generator Form")
        self.width_of_root = 900
        self.height_of_root = 460
        self.screen_width = self.root.winfo_screenwidth()
        self.screen_height = self.root.winfo_screenheight()
        self.x_coordinate = (self.screen_width/2)-(self.width_of_root/2)
        self.y_coordinate = (self.screen_height/2)-(self.height_of_root/1.8)
        self.root.geometry("%dx%d+%d+%d" %(self.width_of_root,self.height_of_root,self.x_coordinate,self.y_coordinate))
        self.root.resizable(width =False, height= False)

    def widgets(self):
        self.frame = Frame(self.root)
        self.frame.pack(padx=20, pady=10)

        self.first_name_label = ttk.Label(self.frame, text="First Name")
        self.first_name_label.grid(row=0, column=0)
        self.last_name_label = ttk.Label(self.frame, text="Last Name")
        self.last_name_label.grid(row=0, column=1)

        self.first_name_entry = ttk.Entry(self.frame, width=30)
        self.last_name_entry = ttk.Entry(self.frame, width=30)
        self.first_name_entry.grid(row=1, column=0)
        self.last_name_entry.grid(row=1, column=1)

        self.phone_label = ttk.Label(self.frame, text="Phone")
        self.phone_label.grid(row=0, column=2)
        self.phone_entry = ttk.Entry(self.frame, width=30)
        self.phone_entry.grid(row=1, column=2)

        self.qty_label = ttk.Label(self.frame, text="Quantity")
        self.qty_label.grid(row=2, column=0)
        self.qty_spinbox = ttk.Spinbox(self.frame, from_=1, to=100, width=28)
        self.qty_spinbox.grid(row=3, column=0)

        self.desc_label = ttk.Label(self.frame, text="Description")
        self.desc_label.grid(row=2, column=1)
        self.desc_entry = ttk.Entry(self.frame, width=30)
        self.desc_entry.grid(row=3, column=1)

        self.price_label = ttk.Label(self.frame, text="Unit Price")
        self.price_label.grid(row=2, column=2)
        self.price_spinbox = ttk.Spinbox(self.frame, from_=0.0, to=500, increment=0.5, width=30)
        self.price_spinbox.grid(row=3, column=2)

        self.add_item_button = ttk.Button(self.frame, text = "Add item", command = self.add_item)
        self.add_item_button.grid(row=4, column=2, pady=5)

        self.columns = ('qty', 'desc', 'price', 'total',)
        self.tree = ttk.Treeview(self.frame, columns=self.columns, show="headings",)
        self.tree.heading('qty', text='Quantity',)
        self.tree.heading('desc', text='Description')
        self.tree.heading('price', text='Unit Price')
        self.tree.heading('total', text="Total")
            
        self.tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10 )

        self.save_invoice_button = ttk.Button(self.frame, text="Generate Invoice", command=self.generate_invoice)
        self.save_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
        self.new_invoice_button = ttk.Button(self.frame, text="New Invoice", command=self.new_invoice)
        self.new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)

    def clear_item(self):
        self.qty_spinbox.delete(0, tkinter.END)
        self.qty_spinbox.insert(0, "1")
        self.desc_entry.delete(0, tkinter.END)
        self.price_spinbox.delete(0, tkinter.END)
        self.price_spinbox.insert(0, "0.0")

        self.invoice_list = []

    def add_item(self):
        self.qty = int(self.qty_spinbox.get())
        self.desc = self.desc_entry.get()
        self.price = float(self.price_spinbox.get())
        self.line_total = self.qty*self.price
        self.invoice_item = [self.qty, self.desc, self.price, self.line_total]
        self.tree.insert('',0, values=self.invoice_item)
        self.clear_item()
        
        self.invoice_list.append(self.invoice_item)

    def new_invoice(self):
        self.first_name_entry.delete(0, tkinter.END)
        self.last_name_entry.delete(0, tkinter.END)
        self.phone_entry.delete(0, tkinter.END)
        self.clear_item()
        self.tree.delete(self.tree.get_children())
        
        self.invoice_list.clear()
        
    def generate_invoice(self):
        self.doc = DocxTemplate("invoice_template.docx")
        self.name = self.first_name_entry.get()+self.last_name_entry.get()
        self.phone = self.phone_entry.get()
        self.subtotal = sum(item[3] for item in self.invoice_list) 
        self.salestax = 0.1
        self.total = self.subtotal*(1-self.salestax)
        
        self.doc.render({"name":self.name, 
                "phone":self.phone,
                "invoice_list": self.invoice_list,
                "subtotal":self.subtotal,
                "salestax":str(self.salestax*100)+"%",
                "total":self.total})
        
        self.doc_name = "new_invoice" + self.name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".docx"
        self.doc.save(self.doc_name)
        
        messagebox.showinfo("Invoice Complete", "Invoice Complete")
        
        self.new_invoice()

if __name__ == "__main__":
    ig = InvoiceGenerator()
    ig.widgets()
    ig.root.mainloop()