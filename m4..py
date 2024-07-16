import matplotlib.pyplot as plt
import numpy as np

# Define the range for y
x = np.linspace(-3, 3, 400)

# Compute u and v
u = x**2 - 1
v = 2 * x

# Plot the results
plt.figure(figsize=(8, 8))
plt.plot(u, v, label='w = z^2 for x = 1')
plt.scatter([1], [0], color='red')  # Point for y = 0
plt.scatter([0], [2], color='blue')  # Point for y = 1
plt.scatter([0], [-2], color='green')  # Point for y = -1
plt.scatter([-3], [4], color='orange')  # Point for y = 2
plt.scatter([-3], [-4], color='purple')  # Point for y = -2

plt.xlabel('Re(w)')
plt.ylabel('Im(w)')
plt.title('Transformation of x = 1 in the z-plane to the w-plane')
plt.axhline(0, color='black',linewidth=0.5)
plt.axvline(0, color='black',linewidth=0.5)
plt.grid(color = 'gray', linestyle = '--', linewidth = 0.5)
plt.legend()
plt.show()
