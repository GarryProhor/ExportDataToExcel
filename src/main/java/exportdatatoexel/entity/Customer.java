package exportdatatoexel.entity;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.FieldDefaults;

@Data
@AllArgsConstructor
@NoArgsConstructor
@Entity
@Table(name = "customers")
@FieldDefaults(level = AccessLevel.PRIVATE)
public class Customer {

    @Id
    @Column(name = "customerNumber")
    Integer customerNumber;

    @Column(name = "customerName")
    String customerName;

    @Column(name = "contactLastName")
    String contactLastName;

    @Column(name = "contactFirstName")
    String contactFirstName;

    @Column(name = "phone")
    String phone;

    @Column(name = "addressLine1")
    String addressLine1;

    @Column(name = "addressLine2")
    String addressLine2;

    @Column(name = "city")
    String city;

    @Column(name = "state")
    String state;

    @Column(name = "postalCode")
    String postalCode;

    @Column(name = "country")
    String country;

    @Column(name = "salesRepEmployeeNumber")
    Integer salesRepEmployeeNumber;

    @Column(name = "creditLimit")
    Long creditLimit;
}
