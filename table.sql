CREATE TABLE Products (
    ProductID INT PRIMARY KEY AUTO_INCREMENT,
    ProductName VARCHAR(50) NOT NULL,
    Description TEXT,
    Price DECIMAL(10, 2) NOT NULL,
    StockQuantity INT NOT NULL CHECK (StockQuantity >= 0),
    DateAdded TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);



INSERT INTO Products (ProductName, Description, Price, StockQuantity) VALUES
('Jacket', 'Warm winter jacket with insulation', 79.99, 50),
('Beanie', 'Soft knit beanie hat for winter', 19.99, 150),
('Gloves', 'Water-resistant winter gloves', 24.99, 100);